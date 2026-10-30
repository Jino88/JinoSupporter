// Pure helpers of the NG Rate pages (ports of the @code blocks of NgRatePage / NgRateAllPage /
// NgRateByGroupPage.razor, NgRateUi.cs and the ngRateTableStyle / chart parts of wwwroot/js/app.js).
import type { HierSubRow, Period, Ppm, ReasonRow } from "../../ngRateApi";

/** Default visible columns of By Model / By Section (the rest behind the Date / Week / Month toggles). */
export const DATE_LIMIT = 14;
export const WEEK_LIMIT = 4;
export const MONTH_LIMIT = 3;
/** By Group shows fewer dates by default. */
export const BY_GROUP_DATE_LIMIT = 10;
/** By Model shows at most this many reason categories. */
export const MAX_REASONS = 10;

export const ppmOf = (ppm: Ppm | undefined, key: string): number => ppm?.[key] ?? 0;

/** "-" for no value, otherwise a rounded thousands-separated number (N0). */
export function formatPpm(value: number): string {
  return value <= 0 ? "-" : Math.round(value).toLocaleString("en-US");
}

export const ppmClass = (value: number): string => (value <= 0 ? "ppm-zero" : "ppm-val");

/** CSS class that colours a process type badge (NgRateUi.ProcessTypeClass). */
export function processTypeClass(processType: string | null | undefined): string {
  const t = (processType ?? "").trim().toUpperCase();
  if (t.startsWith("SUB")) return "pt-sub";
  if (t.startsWith("MAIN")) return "pt-main";
  if (t.startsWith("FUNC")) return "pt-func";
  if (t.includes("VISUAL") || t.startsWith("VIS")) return "pt-vis";
  if (t.startsWith("OQC")) return "pt-oqc";
  return "pt-other";
}

const FALLBACK_BAR_COLORS = ["rgba(249,115,22,0.85)", "rgba(234,179,8,0.85)", "rgba(99,102,241,0.85)", "rgba(236,72,153,0.85)"];

/** Bar colour of a process type in the Summary chart; others cycle through the fallback colours by index. */
export function processTypeColor(processType: string, index: number): string {
  const upper = processType.toUpperCase();
  if (upper.startsWith("SUB")) return "rgba(239,68,68,0.85)";
  if (processType === "MAIN") return "rgba(34,197,94,0.85)";
  if (processType === "FUNCTION") return "rgba(168,85,247,0.85)";
  if (upper.startsWith("VISUAL")) return "rgba(6,182,212,0.85)";
  return FALLBACK_BAR_COLORS[index % FALLBACK_BAR_COLORS.length];
}

export const visibleCols = (cols: Period[], limit: number, expanded: boolean): Period[] =>
  !expanded && cols.length > limit ? cols.slice(0, limit) : cols;

export interface ColumnSet {
  date: Period[];
  week: Period[];
  month: Period[];
}

export const allKeys = (cols: ColumnSet): string[] => [...cols.date, ...cols.week, ...cols.month].map((c) => c.key);

/** True when any of the keys has a PPM above zero. */
export const hasAnyValue = (ppm: Ppm, keys: string[]): boolean => keys.some((k) => ppmOf(ppm, k) > 0);

/** Text colour of a progress-log line (NgRateSetupPanel.GetLogClass). */
export function logClass(line: string): string {
  const lower = line.toLowerCase();
  if (lower.includes("[error]")) return "text-danger";
  if (lower.includes("[warn]")) return "text-warning";
  if (lower.includes("done")) return "text-success";
  if (lower.includes("report complete")) return "text-success";
  return "";
}

/** "[HH:mm:ss] message" as the Blazor pages logged it. */
export function stampLog(message: string, now = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `[${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}] ${message}`;
}

/** yyyy-MM-dd of today (local). */
export function today(now = new Date()): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

/** "2026-10-30" → "10/30/2026" (collapsed setup bar). */
export function usDate(day: string): string {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(day);
  return m ? `${m[2]}/${m[3]}/${m[1]}` : day;
}

/** Copy of the set with the key added or removed. */
export function toggled<T>(set: ReadonlySet<T>, key: T): Set<T> {
  const next = new Set(set);
  if (!next.delete(key)) next.add(key);
  return next;
}

export interface ChartData {
  labels: string[];
  /** One value per label; null at the empty separator labels and where there is no data. */
  values: (ppm: Ppm) => (number | null)[];
}

/**
 * Chart x-axis over the visible columns, newest first per block, with an empty label between the
 * Date / Week / Month blocks (the chart leaves a gap there).
 */
export function chartAxis(cols: ColumnSet): ChartData {
  const blocks = [cols.date, cols.week, cols.month];
  const labels: string[] = [];
  const keys: (string | null)[] = [];
  blocks.forEach((block, i) => {
    if (block.length === 0) return;
    if (i > 0 && keys.length > 0) {
      labels.push("");
      keys.push(null);
    }
    for (const c of block) {
      labels.push(c.header);
      keys.push(c.key);
    }
  });
  return {
    labels,
    values: (ppm) => keys.map((k) => (k === null ? null : ppmOf(ppm, k) > 0 ? ppmOf(ppm, k) : null)),
  };
}

/**
 * Summary-chart x-axis of By Model / By Section: unlike the By Group chart, a separator is added
 * before the Week / Month block even when the Date block is empty.
 */
export function summaryChartAxis(cols: ColumnSet): ChartData {
  const labels: string[] = cols.date.map((c) => c.header);
  const keys: (string | null)[] = cols.date.map((c) => c.key);
  for (const block of [cols.week, cols.month]) {
    if (block.length === 0) continue;
    labels.push("");
    keys.push(null);
    for (const c of block) {
      labels.push(c.header);
      keys.push(c.key);
    }
  }
  return {
    labels,
    values: (ppm) => keys.map((k) => (k === null ? null : ppmOf(ppm, k) > 0 ? ppmOf(ppm, k) : null)),
  };
}

/** Reason rows to show: every Total row, other rows only when they have a PPM; at most `maxReasons` categories. */
export function displayReasonRows(rows: ReasonRow[], maxReasons?: number): ReasonRow[] {
  const shown = rows.filter((r) => r.isTotal || Object.values(r.ppm).some((v) => v > 0));
  if (maxReasons === undefined) return shown;
  const used = new Set<string>();
  const result: ReasonRow[] = [];
  for (const row of shown) {
    if (!used.has(row.reason)) {
      used.add(row.reason);
      if (used.size > maxReasons) break;
    }
    result.push(row);
  }
  return result;
}

/** A sub row shows only while every sub-level ancestor is expanded (HierReportBuilder.IsRowVisible). */
export function isHierRowVisible(row: HierSubRow, byId: ReadonlyMap<string, HierSubRow>, expanded: ReadonlySet<string>): boolean {
  let parentId = row.parentId;
  while (parentId !== null) {
    const parent = byId.get(parentId);
    if (!parent) break;
    if (!expanded.has(parentId)) return false;
    parentId = parent.parentId;
  }
  return true;
}

/** Average of the values above zero; 0 when there is none (TOTAL row of By Group). */
export function meanOfPositive(values: number[]): number {
  const positive = values.filter((v) => v > 0);
  return positive.length === 0 ? 0 : positive.reduce((a, b) => a + b, 0) / positive.length;
}

/** Number in a table cell ("1,234" → 1234; text in parentheses ignored), as ngRateTableStyle.parseNumber. */
export function parseCellNumber(text: string | null | undefined): number | null {
  const match = (text ?? "").replace(/\([^)]*\)/g, "").match(/-?\d[\d,]*(?:\.\d+)?/);
  if (!match) return null;
  const value = Number(match[0].replace(/,/g, ""));
  return Number.isFinite(value) ? value : null;
}

/**
 * Cells holding the column maximum (ties within 0.5 included), as ngRateTableStyle.applyTable:
 * `rows[r][c]` is the cell value or null for cells that do not take part.
 */
export function maxCellPositions(rows: (number | null)[][]): [number, number][] {
  const best = new Map<number, { value: number; cells: [number, number][] }>();
  rows.forEach((cells, r) =>
    cells.forEach((value, c) => {
      if (value === null || value <= 0) return;
      const current = best.get(c);
      if (!current || value > current.value) best.set(c, { value, cells: [[r, c]] });
      else if (Math.abs(value - current.value) < 0.5) current.cells.push([r, c]);
    }),
  );
  return [...best.values()].flatMap((e) => e.cells);
}

/** Selection count label of the Model Groups picker: "3 groups" / "1 group". */
export const plural = (count: number, word: string): string => `${count} ${word}${count === 1 ? "" : "s"}`;
