// Development plan of a Working model (Dev Plan tab): events (ER build, DR, tooling, ...) each with a detailed
// schedule of tasks. Kept in the model information JSON (info.devPlan) and saved with the model.
import { addDays } from "./dates";

export interface PlanTask {
  id: string;
  name: string;
  /** yyyy-MM-dd, or "" when not set. */
  start: string;
  end: string;
  owner: string;
  done: boolean;
  note: string;
  /** Id of the task it is a sub-task of (one level only); absent for a task of the event itself. */
  parentId?: string;
}

/** Group of a quantity row: the materials to get in (from the BOM) or the stages to make (SPK Assy, Function, ...). */
export type QtyGroup = "material" | "making";
export const QTY_GROUPS: { key: QtyGroup; label: string }[] = [
  { key: "material", label: "Material Plan" },
  { key: "making", label: "Making Plan" },
];

/** A row of the quantity table of an event (a material, or SPK Assy, Lower Assy, ...): the quantity planned on each day. */
export interface QtyRow {
  id: string;
  name: string;
  group: QtyGroup;
  /** yyyy-MM-dd -> quantity; days without a quantity are left out. */
  qty: Record<string, number>;
}

export interface PlanEvent {
  id: string;
  name: string;
  /** yyyy-MM-dd; "" = taken from its tasks. */
  start: string;
  end: string;
  color: string;
  tasks: PlanTask[];
  rows: QtyRow[];
  /** Product whose BOM materials were put into the Material Plan once ("" = not yet), so rows deleted stay deleted. */
  bomFrom: string;
}

export interface DevPlan {
  events: PlanEvent[];
}

export const COLORS = ["#4f46e5", "#0891b2", "#059669", "#d97706", "#dc2626", "#7c3aed", "#db2777", "#475569"];

const newId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const DATE = /^\d{4}-\d{2}-\d{2}$/;
const isDate = (s: string) => DATE.test(s);

export const blankEvent = (index: number): PlanEvent =>
  ({ id: newId(), name: "", start: "", end: "", color: COLORS[index % COLORS.length], tasks: [], rows: [], bomFrom: "" });
export const blankRow = (name = "", group: QtyGroup = "making"): QtyRow => ({ id: newId(), name, group, qty: {} });
export const blankTask = (): PlanTask => ({ id: newId(), name: "", start: "", end: "", owner: "", done: false, note: "" });

/** First and last day of something with a start and an end; a missing end is the start and the other way round. */
function ends(x: { start: string; end: string }): [string, string] | null {
  const s = isDate(x.start) ? x.start : isDate(x.end) ? x.end : "";
  const e = isDate(x.end) ? x.end : s;
  return s ? (e < s ? [e, s] : [s, e]) : null;
}

/** The days an event covers: its own dates, else from its first task to its last. */
export function eventSpan(ev: PlanEvent): [string, string] | null {
  const own = ends(ev);
  if (own) return own;
  const spans = ev.tasks.map(ends).filter((x): x is [string, string] => !!x);
  if (!spans.length) return null;
  return [spans.map((x) => x[0]).sort()[0], spans.map((x) => x[1]).sort().at(-1)!];
}

export const taskSpan = (t: PlanTask) => ends(t);

/** Every day from the first to the last of the spans with `pad` days on each side; none without a span. */
export function calendarDays(spans: ([string, string] | null)[], pad = 2): string[] {
  const all = spans.filter((x): x is [string, string] => !!x);
  if (!all.length) return [];
  const from = addDays(all.map((x) => x[0]).sort()[0], -pad);
  const to = addDays(all.map((x) => x[1]).sort().at(-1)!, pad);
  const days: string[] = [];
  for (let d = from; d <= to && days.length < 1500; d = addDays(d, 1)) days.push(d);
  return days;
}

/** The months over the days, for the header: [month "2026-11", number of days]. */
export function monthRuns(days: string[]): [string, number][] {
  const runs: [string, number][] = [];
  for (const d of days) {
    const m = d.slice(0, 7);
    if (runs.length && runs[runs.length - 1][0] === m) runs[runs.length - 1][1]++;
    else runs.push([m, 1]);
  }
  return runs;
}

export const inSpan = (day: string, span: [string, string] | null) => !!span && day >= span[0] && day <= span[1];

/** What is wrong with an event or task before it is kept, or "". */
export function planProblem(x: { name: string; start: string; end: string }): string {
  if (!x.name.trim()) return "Name is required.";
  if (x.start && !isDate(x.start)) return "Start must be a date.";
  if (x.end && !isDate(x.end)) return "End must be a date.";
  if (x.start && x.end && x.end < x.start) return "End is before start.";
  return "";
}

/** The month and day of a yyyy-MM-dd date as "MM/dd" for the date box (the year has its own box beside it). */
export const monthDayText = (date: string) => (isDate(date) ? `${date.slice(5, 7)}/${date.slice(8, 10)}` : date);

/**
 * A month and day typed in the task date box ("11/28", "1/5", "1128", "11-28", "11.28") as a date of that year.
 * Output: yyyy-MM-dd, "" when nothing is typed, null when it is not a day of that year.
 */
export function readMonthDay(text: string, year: number): string | null {
  const t = text.trim();
  if (!t) return "";
  const m = /^(\d{1,2})\s*[/.-]\s*(\d{1,2})$/.exec(t) ?? /^(\d{2})(\d{2})$/.exec(t);
  if (!m) return null;
  const mo = Number(m[1]), d = Number(m[2]);
  const day = new Date(Date.UTC(year, mo - 1, d));
  if (mo < 1 || mo > 12 || d < 1 || day.getUTCMonth() !== mo - 1) return null;
  return `${year}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/**
 * The year and the month/day typed in the two boxes of a date as yyyy-MM-dd.
 * Output: "" when no month/day is typed, null when the year is not 4 digits or the month/day is not a day of that year.
 */
export function readYearMonthDay(year: string, monthDay: string): string | null {
  if (!monthDay.trim()) return "";
  const y = year.trim();
  return /^\d{4}$/.test(y) ? readMonthDay(monthDay, Number(y)) : null;
}

/** The list with one item (dragged) moved to the place of another (dropped on): after it going down, before it going up. */
export function moveTo<T extends { id: string }>(list: T[], fromId: string, toId: string): T[] {
  const i = list.findIndex((x) => x.id === fromId), j = list.findIndex((x) => x.id === toId);
  if (i < 0 || j < 0 || i === j) return list;
  const out = [...list];
  const [item] = out.splice(i, 1);
  out.splice(j, 0, item);
  return out;
}

/**
 * The tasks in the order shown: each task of the event followed by its sub-tasks. A sub-task whose parent is gone, or
 * is itself a sub-task, becomes a task of the event (sub-tasks are one level deep).
 */
export function taskTree(tasks: PlanTask[]): PlanTask[] {
  const tops = new Set(tasks.filter((t) => !t.parentId).map((t) => t.id));
  const fixed: PlanTask[] = tasks.map((t) => (t.parentId && !tops.has(t.parentId) ? (({ parentId: _, ...x }) => x)(t) : t));
  return fixed.filter((t) => !t.parentId).flatMap((t) => [t, ...fixed.filter((c) => c.parentId === t.id)]);
}

/** A new sub-task put right after the last sub-task of its parent. */
export const addSubTask = (tasks: PlanTask[], parentId: string, t: PlanTask): PlanTask[] =>
  taskTree([...tasks, { ...t, parentId }]);

/** A task dragged onto another of the same level: a task moves with its sub-tasks, a sub-task only within its parent. */
export function moveTask(tasks: PlanTask[], fromId: string, toId: string): PlanTask[] {
  const a = tasks.find((t) => t.id === fromId), b = tasks.find((t) => t.id === toId);
  if (!a || !b || (a.parentId ?? "") !== (b.parentId ?? "")) return tasks;
  return taskTree(moveTo(tasks, fromId, toId));
}

/** A task taken out with its sub-tasks. */
export const removeTask = (tasks: PlanTask[], id: string) => tasks.filter((t) => t.id !== id && t.parentId !== id);

/** The days of a task on the calendar: its own, else from its first sub-task to its last. */
export function groupSpan(tasks: PlanTask[], t: PlanTask): [string, string] | null {
  const own = ends(t);
  if (own || t.parentId) return own;
  const spans = tasks.filter((c) => c.parentId === t.id).map(ends).filter((x): x is [string, string] => !!x);
  return spans.length ? [spans.map((x) => x[0]).sort()[0], spans.map((x) => x[1]).sort().at(-1)!] : null;
}

/** Tasks done out of all, for the event row. */
export const progress = (ev: PlanEvent) => [ev.tasks.filter((t) => t.done).length, ev.tasks.length] as const;

/** A quantity typed in a cell: "" (or 0) clears the day, commas are allowed; null when it is not a number. */
export function readQty(text: string): number | null | undefined {
  const t = text.replace(/,/g, "").trim();
  if (!t) return undefined;
  const n = Number(t);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

/** The row with the quantity of a day set, or cleared when undefined or 0; days kept in order. */
export function setQty(row: QtyRow, day: string, n: number | undefined): QtyRow {
  const qty: Record<string, number> = { ...row.qty };
  if (n) qty[day] = n;
  else delete qty[day];
  return { ...row, qty: Object.fromEntries(Object.entries(qty).sort(([a], [b]) => a.localeCompare(b))) };
}

export const rowTotal = (row: QtyRow) => Object.values(row.qty).reduce((a, b) => a + b, 0);

/**
 * The direct materials (level 1) of a BOM the Material Plan of the rows does not have yet, each once, in BOM order.
 * Input: the rows of the event and the BOM rows. Output: the material names to add.
 */
export function bomRowNames(rows: QtyRow[], bom: { level: number; materialName: string; materialCode: string }[]): string[] {
  const k = (s: string) => s.trim().toLowerCase();
  const have = new Set(rows.filter((r) => r.group === "material").map((r) => k(r.name)));
  const out: string[] = [];
  for (const b of bom) {
    const n = (b.materialName || b.materialCode).trim();
    if (b.level !== 1 || !n || have.has(k(n))) continue;
    have.add(k(n));
    out.push(n);
  }
  return out;
}

/** The event with these materials added to its Material Plan and the product its BOM came from kept. */
export const addBomRows = (ev: PlanEvent, names: string[], product: string): PlanEvent =>
  ({ ...ev, bomFrom: product, rows: [...ev.rows, ...names.map((n) => blankRow(n, "material"))] });

/** The rows with one moved up or down past the next row of the same group (rows of the other group stay where they are). */
export function moveInGroup(rows: QtyRow[], id: string, by: -1 | 1): QtyRow[] {
  const i = rows.findIndex((r) => r.id === id);
  if (i < 0) return rows;
  let j = i + by;
  while (j >= 0 && j < rows.length && rows[j].group !== rows[i].group) j += by;
  if (j < 0 || j >= rows.length) return rows;
  const list = [...rows];
  [list[i], list[j]] = [list[j], list[i]];
  return list;
}

/** The days of the plan calendar: the events with 2 days on each side, and every day with a quantity (so none is hidden). */
export function planDays(events: PlanEvent[]): string[] {
  const filled = events.flatMap((e) => e.rows.flatMap((r) => Object.keys(r.qty))).map((d): [string, string] => [d, d]);
  return calendarDays([...events.map(eventSpan), ...filled]);
}

/** A saved plan read leniently (older or hand-edited JSON). */
export function readDevPlan(v: unknown): DevPlan {
  const p = (v && typeof v === "object" ? v : {}) as { events?: unknown };
  const str = (x: unknown) => (typeof x === "string" ? x : "");
  const list = (x: unknown) => (Array.isArray(x) ? x : []) as Record<string, unknown>[];
  return {
    events: list(p.events).filter((e) => e && typeof e === "object").map((e, i) => ({
      id: str(e.id) || newId(),
      name: str(e.name),
      start: str(e.start),
      end: str(e.end),
      color: str(e.color) || COLORS[i % COLORS.length],
      bomFrom: str(e.bomFrom),
      tasks: taskTree(list(e.tasks).filter((t) => t && typeof t === "object").map((t) => ({
        id: str(t.id) || newId(), name: str(t.name), start: str(t.start), end: str(t.end),
        owner: str(t.owner), done: t.done === true, note: str(t.note), ...(str(t.parentId) ? { parentId: str(t.parentId) } : {}),
      }))),
      rows: list(e.rows).filter((r) => r && typeof r === "object").map((r) => ({
        id: str(r.id) || newId(),
        name: str(r.name),
        group: (r.group === "material" ? "material" : "making") as QtyGroup,
        qty: Object.fromEntries(Object.entries((r.qty && typeof r.qty === "object" ? r.qty : {}) as Record<string, unknown>)
          .filter((x): x is [string, number] => isDate(x[0]) && typeof x[1] === "number" && Number.isFinite(x[1]) && x[1] > 0).sort(([a], [b]) => a.localeCompare(b))),
      })),
    })),
  };
}
