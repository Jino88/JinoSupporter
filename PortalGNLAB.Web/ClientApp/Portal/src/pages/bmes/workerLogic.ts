import type { WorkerRecord } from "../../bmesApi";

export const METRICS = ["Day", "Status", "Date Type", "Check In", "Check Out"] as const;
export type Metric = (typeof METRICS)[number];

const cmp = (a: string, b: string) => a.localeCompare(b, undefined, { sensitivity: "base" });
const DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/** yyyy-MM-dd of a record date ("2026-10-30T00:00:00"). */
export const dateKey = (date: string) => date.slice(0, 10);

/** Weekday name of a yyyy-MM-dd date. */
export function weekday(key: string): string {
  const [y, m, d] = key.split("-").map(Number);
  return DAYS[new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
}

/** Minutes after midnight of "HH:mm", "HH:mm:ss" or a date-time text; null when there is no time. */
export function parseTime(value: string | undefined): number | null {
  const m = /(\d{1,2}):(\d{2})(?::(\d{2}))?/.exec((value ?? "").trim());
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]) + (m[3] ? Number(m[3]) / 60 : 0);
}

/** Late when the actual check-in is after the scheduled start. */
export function isLate(r: Pick<WorkerRecord, "checkIn" | "schedStart">): boolean {
  const checkIn = parseTime(r.checkIn);
  const sched = parseTime(r.schedStart);
  return checkIn !== null && sched !== null && checkIn > sched;
}

const mix = (a: number, b: number, t: number) => Math.round(a + (b - a) * t);
const mixRgb = (from: [number, number, number], to: [number, number, number], t: number) =>
  `rgb(${mix(from[0], to[0], t)}, ${mix(from[1], to[1], t)}, ${mix(from[2], to[2], t)})`;

/** Check-in colour: plain bold up to 10 min early, then green → orange over the next 10 minutes. */
export function checkInStyle(r: Pick<WorkerRecord, "checkIn" | "schedStart">): { fontWeight: number; backgroundColor?: string; color?: string } {
  const checkIn = parseTime(r.checkIn);
  const sched = parseTime(r.schedStart);
  if (checkIn === null || sched === null) return { fontWeight: 700 };
  const diff = Math.floor(checkIn - sched);
  if (diff < -10) return { fontWeight: 700 };
  const t = Math.min(Math.max(diff + 10, 0), 10) / 10;
  return { fontWeight: 700, backgroundColor: mixRgb([220, 252, 231], [255, 237, 213], t), color: mixRgb([22, 101, 52], [194, 65, 12], t) };
}

export function statusClass(status: string | undefined): string {
  const s = (status ?? "").toLowerCase();
  if (s.includes("office") || s.includes("출근")) return "ws-in";
  if (s.includes("out") || s.includes("퇴근")) return "ws-out";
  if (s.includes("absent") || s.includes("결근")) return "ws-absent";
  return "ws-other";
}

export interface Summary { rows: number; people: number; days: number; checkedIn: number; checkedOut: number; late: number }

export function summarize(records: WorkerRecord[], totalCount: number): Summary {
  return {
    rows: totalCount,
    people: new Set(records.map((r) => r.empNo)).size,
    days: new Set(records.map((r) => dateKey(r.date))).size,
    checkedIn: records.filter((r) => r.checkIn).length,
    checkedOut: records.filter((r) => r.checkOut).length,
    late: records.filter(isLate).length,
  };
}

export interface Person { key: string; records: WorkerRecord[]; byDate: Map<string, WorkerRecord> }
export interface Dept { key: string; code: string; no: string; name: string; rows: number; people: number; persons: Person[] }

export const deptKey = (code: string, no: string) => `${code}\u001f${no}`;

/** Departments (by code, then no) with their people (by name, then emp no); the first record of a day wins. */
export function buildDepartments(records: WorkerRecord[]): Dept[] {
  const depts = new Map<string, WorkerRecord[]>();
  for (const r of records) {
    const k = deptKey(r.departmentCode, r.departmentNo);
    depts.set(k, [...(depts.get(k) ?? []), r]);
  }
  return [...depts]
    .map(([key, list]) => {
      const people = new Map<string, WorkerRecord[]>();
      for (const r of list) {
        const pk = `${r.empNo}\u001f${r.name}`;
        people.set(pk, [...(people.get(pk) ?? []), r]);
      }
      const persons = [...people]
        .map(([pk, recs]) => {
          const byDate = new Map<string, WorkerRecord>();
          for (const r of recs) if (!byDate.has(dateKey(r.date))) byDate.set(dateKey(r.date), r);
          return { key: pk, records: recs, byDate };
        })
        .sort((a, b) => cmp(a.records[0].name, b.records[0].name) || cmp(a.records[0].empNo, b.records[0].empNo));
      const names = new Set(list.map((r) => (r.empNo.trim() ? r.empNo : r.name).toLowerCase()));
      const first = list[0];
      return { key, code: first.departmentCode, no: first.departmentNo, name: first.departmentName, rows: list.length, people: names.size, persons };
    })
    .sort((a, b) => cmp(a.code, b.code) || cmp(a.no, b.no));
}

/** Date columns, newest first. */
export function dateColumns(records: WorkerRecord[]): string[] {
  return [...new Set(records.map((r) => dateKey(r.date)))].sort().reverse();
}

/** Keeps From <= To after one of them was edited. */
export function normalizeRange(start: string, end: string, edited: "start" | "end"): { start: string; end: string } {
  if (end >= start) return { start, end };
  return edited === "start" ? { start, end: start } : { start: end, end };
}
