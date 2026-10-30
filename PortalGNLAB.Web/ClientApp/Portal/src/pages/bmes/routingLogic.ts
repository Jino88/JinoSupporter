import type { RoutingRow } from "../../bmesApi";

export interface ModelEntry { name: string; count: number }

const cmp = (a: string, b: string) => a.localeCompare(b, undefined, { sensitivity: "base" });
const has = (text: string | undefined, q: string) => (text ?? "").toLowerCase().includes(q.toLowerCase());

/** Models with their process counts, sorted by name. Rows without a model name are left out. */
export function modelCounts(rows: RoutingRow[]): ModelEntry[] {
  const counts = new Map<string, number>();
  for (const r of rows) if (r.modelName) counts.set(r.modelName, (counts.get(r.modelName) ?? 0) + 1);
  return [...counts].map(([name, count]) => ({ name, count })).sort((a, b) => cmp(a.name, b.name));
}

/** Left list filter: a model matches by its name or by any process code / name it has. */
export function filterModels(models: ModelEntry[], rows: RoutingRow[], search: string): ModelEntry[] {
  const q = search.trim();
  if (!q) return models;
  const byProcess = new Set(rows.filter((r) => has(r.processCode, q) || has(r.processName, q)).map((r) => r.modelName));
  return models.filter((m) => has(m.name, q) || byProcess.has(m.name));
}

export function selectedRows(rows: RoutingRow[], model: string, type: string): RoutingRow[] {
  if (!model) return [];
  return rows.filter((r) => r.modelName === model && (!type || r.processType === type));
}

export function selectedTypes(rows: RoutingRow[], model: string): string[] {
  const types = new Set(rows.filter((r) => r.modelName === model && r.processType.trim()).map((r) => r.processType));
  return [...types].sort(cmp);
}

/** Model to show after the rows changed: keep the current one while it has rows, else the first. */
export function pickModel(current: string, rows: RoutingRow[]): string {
  if (current && rows.some((r) => r.modelName === current)) return current;
  return modelCounts(rows)[0]?.name ?? "";
}
