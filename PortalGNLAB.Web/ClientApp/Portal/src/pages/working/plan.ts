// Development plan of a Working model (Dev Plan tab): events (ER build, DR, tooling, ...) each with a detailed
// schedule of tasks. Kept in the model information JSON (info.devPlan) and saved with the model.
import { addDays } from "../productionInput/logic";

export interface PlanTask {
  id: string;
  name: string;
  /** yyyy-MM-dd, or "" when not set. */
  start: string;
  end: string;
  owner: string;
  done: boolean;
  note: string;
}

/** A row of the quantity table of an event (SPK Assy, Lower Assy, ...): the quantity planned on each day. */
export interface QtyRow {
  id: string;
  name: string;
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
}

export interface DevPlan {
  events: PlanEvent[];
}

export const COLORS = ["#4f46e5", "#0891b2", "#059669", "#d97706", "#dc2626", "#7c3aed", "#db2777", "#475569"];

const newId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const DATE = /^\d{4}-\d{2}-\d{2}$/;
const isDate = (s: string) => DATE.test(s);

export const blankEvent = (index: number): PlanEvent =>
  ({ id: newId(), name: "", start: "", end: "", color: COLORS[index % COLORS.length], tasks: [], rows: [] });
export const blankRow = (name = ""): QtyRow => ({ id: newId(), name, qty: {} });
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
      tasks: list(e.tasks).filter((t) => t && typeof t === "object").map((t) => ({
        id: str(t.id) || newId(), name: str(t.name), start: str(t.start), end: str(t.end),
        owner: str(t.owner), done: t.done === true, note: str(t.note),
      })),
      rows: list(e.rows).filter((r) => r && typeof r === "object").map((r) => ({
        id: str(r.id) || newId(),
        name: str(r.name),
        qty: Object.fromEntries(Object.entries((r.qty && typeof r.qty === "object" ? r.qty : {}) as Record<string, unknown>)
          .filter((x): x is [string, number] => isDate(x[0]) && typeof x[1] === "number" && Number.isFinite(x[1]) && x[1] > 0).sort(([a], [b]) => a.localeCompare(b))),
      })),
    })),
  };
}
