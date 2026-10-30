// F-Cost page logic: port of the view helpers in Components/Pages/BmesFCostPage.razor
// (column folding, Mid → Sub1 → Sub2 hierarchy, totals, CSV).
import type { FCostView, PeriodKind } from "../../bmesReportApi";
import { cmpIgnoreCase, cmpOrdinal, distinctIgnoreCase, groupBy, n0 } from "./format";

export const COLLAPSE_LIMIT: Record<PeriodKind, number> = { Day: 8, Week: 5, Month: 4, Unknown: Number.MAX_SAFE_INTEGER };

export interface Metrics { input: boolean; cost: boolean; rate: boolean }

export type Expanded = Record<"Day" | "Week" | "Month", boolean>;

export const valueAt = (values: readonly number[] | null | undefined, index: number) =>
  values && index >= 0 && index < values.length ? values[index] : 0;

export const formatNum = (v: number) => (v === 0 ? "-" : n0(v));
export const formatRate = (v: number) => (v === 0 ? "-" : v.toFixed(2) + "%");
export const rateOf = (input: number, cost: number) => (input > 0 ? (cost / input) * 100 : 0);

/** Start of the default period: the first day of the month two months back. */
export function defaultStartDate(today: Date): Date {
  return new Date(today.getFullYear(), today.getMonth() - 2, 1);
}

export function periodColumnCount(view: FCostView, kind: PeriodKind): number {
  return view.columns.filter((c) => c.kind === kind).length;
}

export function visiblePeriodColumnCount(view: FCostView, kind: "Day" | "Week" | "Month", expanded: Expanded): number {
  const total = periodColumnCount(view, kind);
  const limit = COLLAPSE_LIMIT[kind];
  return expanded[kind] || total <= limit ? total : limit;
}

/** Day / Week / Month blocks show their first N columns until expanded (VisibleFCostColumnIndexes). */
export function visibleColumnIndexes(view: FCostView, expanded: Expanded): number[] {
  if (view.columns.length === 0) return view.headers.map((_, i) => i);
  const totals = { Day: 0, Week: 0, Month: 0 };
  for (const c of view.columns) if (c.kind !== "Unknown") totals[c.kind]++;
  const seen = { Day: 0, Week: 0, Month: 0 };
  const indexes: number[] = [];
  view.columns.forEach((c, i) => {
    if (c.kind === "Unknown") {
      indexes.push(i);
      return;
    }
    const n = ++seen[c.kind];
    const limit = COLLAPSE_LIMIT[c.kind];
    if (totals[c.kind] <= limit || expanded[c.kind] || n <= limit) indexes.push(i);
  });
  return indexes;
}

export function allColumnIndexes(view: FCostView): number[] {
  const count = view.columns.length > 0 ? view.columns.length : view.headers.length;
  return Array.from({ length: count }, (_, i) => i);
}

/** Heavier divider where the period kind changes (ColBoundaryClass). */
export function colBoundaryClass(view: FCostView, idx: number): string {
  const cols = view.columns;
  if (idx <= 0 || idx >= cols.length) return "";
  const cur = cols[idx].kind;
  if (cur === cols[idx - 1].kind) return "";
  return cur === "Day" ? "col-day-start" : cur === "Week" ? "col-week-start" : cur === "Month" ? "col-month-start" : "";
}

const hasInputAndCost = (input: number[], cost: number[]) => input.some((v) => v > 0) && cost.some((v) => v > 0);

/** Report totals for the given columns: BMES' own Total when it has both Input and F-COST, else the
 *  sum of the materials, else the sum of the sub groups (ComputeSubGroupTotals). */
export function computeTotals(view: FCostView, columnIndexes: number[]): { input: number[]; cost: number[] } {
  const n = columnIndexes.length;
  if (view.total) {
    const input = columnIndexes.map((ci) => valueAt(view.total!.input, ci));
    const cost = columnIndexes.map((ci) => valueAt(view.total!.cost, ci));
    if (hasInputAndCost(input, cost)) return { input, cost };
  }
  let input = new Array<number>(n).fill(0);
  let cost = new Array<number>(n).fill(0);
  for (const m of view.materials) {
    columnIndexes.forEach((ci, i) => {
      input[i] += valueAt(m.input, ci);
      cost[i] += valueAt(m.cost, ci);
    });
  }
  if (hasInputAndCost(input, cost)) return { input, cost };
  input = new Array<number>(n).fill(0);
  cost = new Array<number>(n).fill(0);
  for (const sg of view.subGroups) {
    columnIndexes.forEach((ci, i) => {
      if (ci < 0) return;
      input[i] += valueAt(sg.inputByCol, ci);
      cost[i] += valueAt(sg.costByCol, ci);
    });
  }
  return { input, cost };
}

export interface HierarchyRow {
  rowId: string;
  parentId: string | null;
  display: string;
  depth: number;
  inputByCol: number[];
  costByCol: number[];
}

/** Path part of a SubGroupKey ("Group::Material::path"), split on " / ". */
export function parseSubPathSegments(subGroupKey: string): string[] {
  const first = subGroupKey.indexOf("::");
  if (first < 0) return [];
  const second = subGroupKey.indexOf("::", first + 2);
  if (second < 0) return [];
  const path = subGroupKey.slice(second + 2);
  return path ? path.split(" / ") : [];
}

/** Mid → Sub1 → Sub2 rows (BuildFCostHierarchy). Parents carry the sums of their leaves. */
export function buildHierarchy(view: FCostView): HierarchyRow[] {
  const rows: HierarchyRow[] = [];
  if (view.subGroups.length === 0) return rows;
  const colCount = view.headers.length;
  const sum = (list: { inputByCol: number[]; costByCol: number[] }[]) => {
    const input = new Array<number>(colCount).fill(0);
    const cost = new Array<number>(colCount).fill(0);
    for (const sg of list) {
      for (let i = 0; i < colCount; i++) {
        input[i] += valueAt(sg.inputByCol, i);
        cost[i] += valueAt(sg.costByCol, i);
      }
    }
    return { input, cost };
  };

  const byMid = [...groupBy(view.subGroups, (sg) => sg.groupName + "\u0000" + sg.midGroupName).values()].sort(
    (a, b) => cmpOrdinal(a[0].groupName, b[0].groupName) || cmpOrdinal(a[0].midGroupName, b[0].midGroupName),
  );

  for (const mid of byMid) {
    const midRowId = `M::${mid[0].groupName}::${mid[0].midGroupName}`;
    const midSum = sum(mid);
    rows.push({ rowId: midRowId, parentId: null, display: mid[0].midGroupName, depth: 0, inputByCol: midSum.input, costByCol: midSum.cost });

    const items = mid.map((sg) => ({ sg, segments: parseSubPathSegments(sg.subGroupKey) }));
    const bySub1 = groupBy(items, (x) => (x.segments.length > 0 ? x.segments[0] : x.sg.display));
    for (const [sub1Display, sub1] of bySub1) {
      const sub1RowId = `S1::${midRowId}::${sub1Display}`;
      const s1 = sum(sub1.map((x) => x.sg));
      rows.push({ rowId: sub1RowId, parentId: midRowId, display: sub1Display, depth: 1, inputByCol: s1.input, costByCol: s1.cost });
      for (const x of sub1) {
        if (x.segments.length <= 1) continue;
        rows.push({
          rowId: `S2::${x.sg.subGroupKey}`,
          parentId: sub1RowId,
          display: x.segments[x.segments.length - 1],
          depth: 2,
          inputByCol: x.sg.inputByCol,
          costByCol: x.sg.costByCol,
        });
      }
    }
  }
  return rows;
}

/** A row shows when every ancestor is expanded. */
export function isRowVisible(row: HierarchyRow, byId: Map<string, HierarchyRow>, expanded: ReadonlySet<string>): boolean {
  let pid = row.parentId;
  while (pid !== null) {
    if (!expanded.has(pid)) return false;
    pid = byId.get(pid)?.parentId ?? null;
  }
  return true;
}

export const levelLabel = (row: HierarchyRow) => (row.depth === 0 ? "Model" : "Sub Group");

export const metricCount = (m: Metrics) => (m.input ? 1 : 0) + (m.cost ? 1 : 0) + (m.rate ? 1 : 0);

/** Link that opens Model Group focused on the unmapped materials (BuildUnmappedModelGroupUrl). */
export function unmappedModelGroupUrl(unmapped: { material: string }[]): string {
  const materials = distinctIgnoreCase(unmapped.map((m) => m.material).filter((m) => m && m.trim())).sort(cmpIgnoreCase);
  if (materials.length === 0) return "/bmes/make-model-group";
  return "/bmes/make-model-group?source=fcost&" + materials.map((m) => "m=" + encodeURIComponent(m)).join("&");
}

/** Rows of the "Export CSV" file of the F-Cost page (BuildRegularFCostCsvRows over every column). */
export function buildCsvRows(view: FCostView, metrics: Metrics): string[][] {
  const columns = allColumnIndexes(view);
  if (columns.length === 0) return [];
  const header = ["Section", "Level", "Name", "Type", ...columns.map((ci) => view.headers[ci] ?? "")];
  const metricRows = (section: string, level: string, name: string, inputOf: (ci: number) => number, costOf: (ci: number) => number) => {
    const out: string[][] = [];
    if (metrics.input) out.push([section, level, name, "INPUT", ...columns.map((ci) => formatNum(inputOf(ci)))]);
    if (metrics.cost) out.push([section, level, name, "F-COST", ...columns.map((ci) => formatNum(costOf(ci)))]);
    if (metrics.rate) out.push([section, level, name, "RATE %", ...columns.map((ci) => formatRate(rateOf(inputOf(ci), costOf(ci))))]);
    return out;
  };

  const totals = computeTotals(view, columns);
  const totalInput = (ci: number) => totals.input[columns.indexOf(ci)] ?? 0;
  const totalCost = (ci: number) => totals.cost[columns.indexOf(ci)] ?? 0;

  const rows: string[][] = [["F-Cost Report", `${view.startDate} - ${view.endDate}`], ["Model / Sub Group Summary"], header];
  rows.push(...metricRows("Model / Sub Group Summary", "Total", "Total", totalInput, totalCost));
  for (const row of buildHierarchy(view)) {
    rows.push(...metricRows("Model / Sub Group Summary", levelLabel(row), row.display,
      (ci) => valueAt(row.inputByCol, ci), (ci) => valueAt(row.costByCol, ci)));
  }
  rows.push([], ["Material Detail"], header);
  if (view.subGroups.length > 0) {
    rows.push(...metricRows("Material Detail", "Total", view.total?.displayName || "Total", totalInput, totalCost));
  }
  for (const m of view.materials) {
    rows.push(...metricRows("Material Detail", m.productGroup, m.displayName, (ci) => valueAt(m.input, ci), (ci) => valueAt(m.cost, ci)));
  }
  return rows;
}
