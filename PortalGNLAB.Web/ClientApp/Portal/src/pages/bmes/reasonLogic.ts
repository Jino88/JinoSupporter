import type { Pair, ReasonRow } from "../../bmesApi";

export const MISSING_KEY = "\u0001missing";
export const OTHER_KEY = "\u0001other";
export const PAGE_SIZE = 300;

export interface Chip { key: string; label: string; count: number }
export interface DisplayRow { processName: string; ngName: string; row: ReasonRow | null }

const cmp = (a: string, b: string) => a.localeCompare(b, undefined, { sensitivity: "base" });
const lower = (s: string | undefined) => (s ?? "").trim().toLowerCase();
const empty = (s: string | undefined) => !(s ?? "").trim();

/** Distinct non-empty reasons, case-insensitive, sorted (the suggestions of the reason inputs). */
export function reasonOptions(rows: ReasonRow[]): string[] {
  const seen = new Map<string, string>();
  for (const r of rows) {
    const v = (r.reason ?? "").trim();
    if (v && !seen.has(v.toLowerCase())) seen.set(v.toLowerCase(), v);
  }
  return [...seen.values()].sort(cmp);
}

/** All / Missing reason / one chip per reason, most used first. */
export function chips(rows: ReasonRow[]): Chip[] {
  const groups = new Map<string, { label: string; count: number }>();
  for (const r of rows) {
    if (empty(r.reason)) continue;
    const key = r.reason.trim().toLowerCase();
    const g = groups.get(key);
    if (g) g.count++;
    else groups.set(key, { label: r.reason.trim(), count: 1 });
  }
  const reasons = [...groups.values()]
    .sort((a, b) => b.count - a.count || cmp(a.label, b.label))
    .map((g) => ({ key: g.label, label: g.label, count: g.count }));
  return [
    { key: "", label: "All", count: rows.length },
    { key: MISSING_KEY, label: "Missing reason", count: rows.filter((r) => empty(r.reason)).length },
    ...reasons,
  ];
}

export function matchesChip(chip: string, reason: string | undefined): boolean {
  if (!chip) return true;
  if (chip === MISSING_KEY) return empty(reason);
  return lower(reason) === chip.trim().toLowerCase();
}

export function matchesSearch(search: string, processName: string, ngName: string, reason: string | undefined): boolean {
  const q = search.trim().toLowerCase();
  if (!q) return true;
  return [processName, ngName, reason ?? ""].some((v) => v.toLowerCase().includes(q));
}

export const pairKey = (processName: string, ngName: string) => `${processName}\u0000${ngName}`;

/** First row wins when a pair exists twice (same as the Blazor page). */
export function pairIndex(rows: ReasonRow[]): Map<string, ReasonRow> {
  const index = new Map<string, ReasonRow>();
  for (const r of rows) {
    const key = pairKey(r.processName ?? "", r.ngName ?? "");
    if (!index.has(key)) index.set(key, r);
  }
  return index;
}

export interface DisplayInput {
  rows: ReasonRow[];
  chip: string;
  search: string;
  editId: number;
  /** null = table view; otherwise the (ProcessName, NgName) pairs found for the typed model. */
  modelPairs: Pair[] | null;
}

/** Rows shown: the table itself, or the model's pairs (red when not in the table yet). */
export function displayRows({ rows, chip, search, editId, modelPairs }: DisplayInput): DisplayRow[] {
  if (modelPairs) {
    const index = pairIndex(rows);
    return modelPairs
      .map((p) => ({ processName: p.processName, ngName: p.ngName, row: index.get(pairKey(p.processName, p.ngName)) ?? null }))
      .filter((d) => (d.row ? matchesChip(chip, d.row.reason) : !chip || chip === MISSING_KEY))
      .filter((d) => matchesSearch(search, d.processName, d.ngName, d.row?.reason));
  }
  return rows
    .filter((r) => r.id === editId || (matchesChip(chip, r.reason) && matchesSearch(search, r.processName ?? "", r.ngName ?? "", r.reason)))
    .map((r, i) => ({ r, i }))
    .sort((a, b) => Number(b.r.id === editId) - Number(a.r.id === editId) || a.i - b.i)
    .map(({ r }) => ({ processName: r.processName ?? "", ngName: r.ngName ?? "", row: r }));
}

/** Options of the quick-edit select: the known reasons plus the current value when it is new. */
export function quickOptions(options: string[], current: string): string[] {
  const cur = current.trim();
  return cur && !options.some((o) => o.toLowerCase() === cur.toLowerCase()) ? [cur, ...options] : options;
}

/** Status text after "Find new pairs". */
export function findPairsMessage(r: { apiAdded: number; cacheAdded: number; apiFailure: string | null; cacheFailure: string | null }): { text: string; error: boolean; added: boolean } {
  const total = r.apiAdded + r.cacheAdded;
  let text = total > 0
    ? `Added ${total} new pair(s) (API ${r.apiAdded}, cache ${r.cacheAdded}).`
    : "No new pairs. The Reason Table already covers all known values.";
  const problems = [r.apiFailure, r.cacheFailure].filter((s): s is string => !!s);
  if (problems.length) text += " " + problems.join("; ") + ".";
  if (total > 0) text += " Click Update To Server to sync.";
  return { text, error: problems.length > 0, added: total > 0 };
}
