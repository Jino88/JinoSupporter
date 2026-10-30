import type { GaussPoint, QrCell, QrData, QrDataRow, QrSummaryRow } from "../../qrApi";
import { CHROME_HIDDEN_QUERY } from "../../menu";

// Pure logic of the QR pages, ported from QrSummaryPage.razor, QrTestDataBrowser.razor and
// QrGaussScatterPage.razor. Nothing here touches the DOM or the network.

/** Sentinel of QrModelCodeService.UnmappedCode: "a QR whose code is not in the table". */
export const UNMAPPED_CODE = "#other";

export const MAX_ROW_CHOICES = [500, 2000, 5000, 10000, 20000];
export const DATA_MAX_ROW_CHOICES = [500, 1000, 2000, 5000, 10000, 20000];

/** The offered caps with whatever is in force folded in, so a value that is not a preset still shows. */
export function maxRowOptions(choices: number[], current: number): number[] {
  return [...new Set([...choices, current].filter((v) => v > 0))].sort((a, b) => a - b);
}

export const num0 = (value: number) => Math.round(value).toLocaleString("en-US");

export function formatPpm(value: number | null | undefined): string {
  return value == null ? "-" : `${num0(value)} ppm`;
}

/** Model code inside a QR (characters 11-12), or "" when the QR is too short to carry one. */
export function extractModelCode(productId: string | null | undefined): string {
  const value = (productId ?? "").trim();
  return value.length < 12 ? "" : value.substring(10, 12).toUpperCase();
}

export interface ModelOption {
  code: string;
  label: string;
}

export interface ModelOptions {
  options: ModelOption[];
  known: Set<string>;
}

/**
 * The model dropdown: the settings table, keeping only the models the fetched QR actually contain
 * and counting each, plus "Other" for QR the table does not name.
 */
export function buildModelOptions(codes: { code: string; name: string }[], productIds: string[]): ModelOptions {
  const known = new Set<string>();
  const names = new Map<string, string>();
  for (const entry of codes) {
    const code = entry.code.toUpperCase();
    known.add(code);
    names.set(code, entry.name.trim() === "" ? entry.code : entry.name);
  }

  const counts = new Map<string, number>();
  let unmapped = 0;
  for (const id of productIds) {
    const code = extractModelCode(id);
    if (known.has(code)) counts.set(code, (counts.get(code) ?? 0) + 1);
    else unmapped++;
  }

  const options: ModelOption[] = [];
  for (const [code, name] of [...names.entries()].sort((a, b) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0))) {
    const count = counts.get(code);
    if (count !== undefined) options.push({ code, label: `${name} (${code}) · ${num0(count)} QR` });
  }
  if (unmapped > 0) options.push({ code: UNMAPPED_CODE, label: `Other · ${num0(unmapped)} QR` });
  return { options, known };
}

/** A model that is no longer on the list must not keep filtering invisibly. */
export function validModel(modelCode: string, options: ModelOption[]): string {
  return modelCode !== "" && !options.some((o) => o.code.toLowerCase() === modelCode.toLowerCase()) ? "" : modelCode;
}

export function matchesModel(productId: string, modelCode: string, known: Set<string>): boolean {
  if (modelCode === "") return true;
  const code = extractModelCode(productId);
  return modelCode === UNMAPPED_CODE ? !known.has(code) : code.toLowerCase() === modelCode.toLowerCase();
}

// ---------------------------------------------------------------- QR SUMMARY

export type RowFilter = "all" | "ng" | "complete";

export interface SummaryLine {
  row: QrSummaryRow;
  isChild: boolean;
  expanded: boolean;
}

export function filterSummaryRows(rows: QrSummaryRow[], modelCode: string, known: Set<string>, filter: RowFilter): QrSummaryRow[] {
  return rows
    .filter((r) => matchesModel(r.productId, modelCode, known))
    .filter((r) => (filter === "ng" ? r.hasAnyFail : filter === "complete" ? r.complete : true));
}

/** One entry per rendered row: expanded groups contribute their earlier rounds in place, newest first. */
export function buildSummaryLines(rows: QrSummaryRow[], expanded: Set<string>): SummaryLine[] {
  const lines: SummaryLine[] = [];
  for (const row of rows) {
    const open = row.older.length > 0 && expanded.has(row.productId);
    lines.push({ row, isChild: false, expanded: open });
    if (open) for (const older of row.older) lines.push({ row: older, isChild: true, expanded: false });
  }
  return lines;
}

export interface StationCard {
  station: string;
  tested: number;
  ng: number;
  ngRatePpm: number | null;
}

export function stationCards(rows: QrSummaryRow[]): StationCard[] {
  const build = (station: string, pick: (r: QrSummaryRow) => QrCell | null): StationCard => {
    let tested = 0;
    let ng = 0;
    for (const row of rows) {
      const cell = pick(row);
      if (!cell) continue;
      tested++;
      if (cell.verdict === "fail") ng++;
    }
    return { station, tested, ng, ngRatePpm: tested <= 0 ? null : (ng / tested) * 1_000_000 };
  };
  return [build("AIRLEAK", (r) => r.airleak), build("BAKO", (r) => r.bako), build("GAUSS", (r) => r.gauss)];
}

export function itemText(cell: QrCell | null, index: number): string {
  const value = cell?.items[index]?.value ?? "";
  return value.trim() === "" ? "-" : value;
}

export const itemFail = (cell: QrCell | null, index: number) => cell?.items[index]?.verdict === "fail";

/** MM-dd HH:mm:ss out of a "yyyy-MM-dd HH:mm:ss" stamp. */
export const formatStamp = (time: string | null | undefined) => (time ? time.slice(5) : "-");

/** GSTD stores 1/0, ALTD stores OK/AL — normalised so one column reads consistently. */
export function verdictLabel(cell: QrCell | null): string {
  if (!cell) return "-";
  if (cell.verdict === "pass") return "OK";
  if (cell.verdict === "fail") return "NG";
  return cell.result.trim() === "" ? "?" : cell.result;
}

export function verdictClass(cell: QrCell | null): string {
  if (!cell) return "qr-none";
  return cell.verdict === "pass" ? "qr-pass" : cell.verdict === "fail" ? "qr-fail" : "";
}

/** The data browser that holds a station's raw rows. */
export function stationRoute(station: string): string {
  switch (station) {
    case "AIRLEAK": return "/app/qr/airleak-data";
    case "BAKO": return "/app/qr/bako-data";
    case "GAUSS": return "/app/qr/gauss-data";
    default: return "";
  }
}

/** Opens the station's data page on one QR; the station's test date is preselected. */
export function stationDetailUrl(productId: string, cell: QrCell): string {
  const route = stationRoute(cell.station);
  if (route === "" || productId.trim() === "") return "";
  let url = `${route}?qr=${encodeURIComponent(productId)}`;
  if (cell.time) url += `&date=${cell.time.slice(0, 10)}`;
  return url + CHROME_HIDDEN_QUERY;
}

export function scatterUrl(from: string, to: string, max: number): string {
  return `/app/qr/gauss-scatter?from=${from}&to=${to}&max=${max}${CHROME_HIDDEN_QUERY}`;
}

export function addDays(date: string, days: number): string {
  const d = new Date(date + "T00:00:00Z");
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

export const localToday = () => {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
};

export function compactDate(date: string) {
  return date.replace(/-/g, "");
}

export function clockStamp(date = new Date()) {
  const p = (n: number) => String(n).padStart(2, "0");
  return `${p(date.getHours())}${p(date.getMinutes())}${p(date.getSeconds())}`;
}

// ---------------------------------------------------------------- data browser

/** Loose match: case and punctuation are ignored, so "Product_ID" finds ProductID. */
const normalizeName = (name: string) => name.replace(/[^\p{L}\p{N}]/gu, "").toLowerCase();

export function findColumnIndex(columns: { name: string }[], expected: string): number {
  const wanted = normalizeName(expected);
  return columns.findIndex((c) => normalizeName(c.name) === wanted);
}

/** Index of the derived test-date column inside QrData.layout and in the filter map. */
export const TEST_DATE_INDEX = -1;

export interface RowGroup {
  key: string;
  rows: QrDataRow[];
}

export const cellValue = (row: QrDataRow, index: number) => (index >= 0 && index < row.values.length ? row.values[index] : "");

/** One group per ProductID (newest row first, as the server returns them); rows with no ProductID stand alone. */
export function buildRowGroups(data: QrData): RowGroup[] {
  const productIdIndex = findColumnIndex(data.columns, "ProductID");
  if (productIdIndex < 0) return data.rows.map((row, i) => ({ key: `row:${i}`, rows: [row] }));

  const groups: RowGroup[] = [];
  const byProduct = new Map<string, RowGroup>();
  data.rows.forEach((row, i) => {
    const productId = (productIdIndex < row.values.length ? row.values[productIdIndex] : "").trim();
    if (productId === "") {
      groups.push({ key: `row:${i}`, rows: [row] });
      return;
    }
    const lookup = productId.toLowerCase();
    let group = byProduct.get(lookup);
    if (!group) {
      group = { key: `product:${productId}`, rows: [] };
      byProduct.set(lookup, group);
      groups.push(group);
    }
    group.rows.push(row);
  });
  return groups;
}

/** Per-item verdict columns: every `<item>_Result` except the unit's own TestResult. */
export function buildNgItemColumns(columns: { name: string }[]): { index: number; label: string }[] {
  const items: { index: number; label: string }[] = [];
  columns.forEach((column, index) => {
    const name = column.name;
    if (name.toLowerCase() === "testresult" || !name.toLowerCase().endsWith("_result")) return;
    items.push({ index, label: name.slice(0, -"_Result".length) });
  });
  return items;
}

export interface Judge {
  isPass: (value: string) => boolean;
  isFail: (value: string) => boolean;
}

/** GSTD spells its verdict "1"/"0" and ALTD spells a leak "AL", so the words come from the table definition. */
export function makeJudge(passValues: string[], failValues: string[]): Judge {
  const has = (list: string[], value: string) => list.some((v) => v.toLowerCase() === value.trim().toLowerCase());
  return { isPass: (v) => has(passValues, v), isFail: (v) => has(failValues, v) };
}

export interface DataFilters {
  /** Matches any cell; set by a QR SUMMARY deep link. */
  search: string;
  modelCode: string;
  /** Per-column text, keyed by column index (TEST_DATE_INDEX for the derived date). */
  columns: Record<number, string>;
}

export const isFiltered = (f: DataFilters) =>
  f.search.trim() !== "" || f.modelCode !== "" || Object.values(f.columns).some((v) => v.trim() !== "");

const includesCi = (hay: string, needle: string) => hay.toLowerCase().includes(needle.toLowerCase());

export function rowMatches(row: QrDataRow, f: DataFilters, productIdIndex: number, known: Set<string>): boolean {
  if (f.modelCode !== "") {
    if (productIdIndex < 0) return false;
    if (!matchesModel(cellValue(row, productIdIndex), f.modelCode, known)) return false;
  }
  const term = f.search.trim();
  if (term !== "" && !row.values.some((v) => includesCi(v, term))) return false;
  for (const [key, text] of Object.entries(f.columns)) {
    const needle = text.trim();
    if (needle === "") continue;
    const index = Number(key);
    const cell = index === TEST_DATE_INDEX ? (row.date ?? "-") : cellValue(row, index);
    if (!includesCi(cell, needle)) return false;
  }
  return true;
}

/**
 * Filters row by row and regroups what survives, so TestResult = FAIL leaves only the failing
 * rounds rather than every round of a QR that failed once. Group keys carry over unchanged.
 */
export function filterGroups(groups: RowGroup[], f: DataFilters, productIdIndex: number, known: Set<string>): RowGroup[] {
  if (!isFiltered(f)) return groups;
  const result: RowGroup[] = [];
  for (const group of groups) {
    const kept = group.rows.filter((row) => rowMatches(row, f, productIdIndex, known));
    if (kept.length > 0) result.push({ key: group.key, rows: kept });
  }
  return result;
}

export interface DataLine {
  group: RowGroup;
  row: QrDataRow;
  isChild: boolean;
  expanded: boolean;
}

export function buildDataLines(groups: RowGroup[], expanded: Set<string>): DataLine[] {
  const lines: DataLine[] = [];
  for (const group of groups) {
    const open = group.rows.length > 1 && expanded.has(group.key);
    lines.push({ group, row: group.rows[0], isChild: false, expanded: open });
    if (open) for (let i = 1; i < group.rows.length; i++) lines.push({ group, row: group.rows[i], isChild: true, expanded: true });
  }
  return lines;
}

/** Input / NG over the QR left on screen: one NG per QR whose newest surviving row is a fail. */
export function filteredSummary(groups: RowGroup[], testResultIndex: number, judge: Judge) {
  let ng = 0;
  if (testResultIndex >= 0) for (const g of groups) if (judge.isFail(cellValue(g.rows[0], testResultIndex))) ng++;
  const input = groups.length;
  return { inputCount: input, ngCount: ng, ngRatePpm: input <= 0 ? null : (ng / input) * 1_000_000 };
}

export interface NgItemGroup {
  label: string;
  total: number;
  children: { label: string; count: number }[];
}

export interface NgBreakdown {
  groups: NgItemGroup[];
  comboCount: number;
  itemTotalSum: number;
  rowCount: number;
  multiCount: number;
  unspecified: number;
}

/**
 * NG on screen as item totals, each split by what the item failed together with. A QR that trips
 * three items is counted under all three, so the totals sum higher than the NG QR count while the
 * children of any one item add up to its own total.
 */
export function buildNgBreakdown(
  groups: RowGroup[],
  itemColumns: { index: number; label: string }[],
  testResultIndex: number,
  judge: Judge,
): NgBreakdown {
  const empty: NgBreakdown = { groups: [], comboCount: 0, itemTotalSum: 0, rowCount: 0, multiCount: 0, unspecified: 0 };
  if (itemColumns.length === 0 || testResultIndex < 0) return empty;

  const splits = itemColumns.map(() => new Map<string, number>());
  const totals = itemColumns.map(() => 0);
  const combos = new Set<string>();
  let rowCount = 0;
  let multiCount = 0;
  let unspecified = 0;

  for (const group of groups) {
    const row = group.rows[0];
    if (!judge.isFail(cellValue(row, testResultIndex))) continue;
    rowCount++;
    const failed: number[] = [];
    itemColumns.forEach((column, i) => {
      if (judge.isFail(cellValue(row, column.index))) failed.push(i);
    });
    if (failed.length === 0) {
      unspecified++;
      continue;
    }
    if (failed.length > 1) multiCount++;
    combos.add(failed.map((i) => itemColumns[i].label).join(" + "));
    for (const i of failed) {
      totals[i]++;
      const others = failed.filter((o) => o !== i).map((o) => itemColumns[o].label).join(" + ");
      splits[i].set(others, (splits[i].get(others) ?? 0) + 1);
    }
  }

  const cmp = (a: string, b: string) => (a < b ? -1 : a > b ? 1 : 0);
  const result = itemColumns
    .map((column, i): NgItemGroup => ({
      label: column.label,
      total: totals[i],
      children: [...splits[i].entries()]
        .sort((a, b) => (a[0] === "" ? 0 : 1) - (b[0] === "" ? 0 : 1) || b[1] - a[1] || cmp(a[0], b[0]))
        .map(([key, count]) => ({ label: key === "" ? "alone" : "+ " + key, count })),
    }))
    .filter((g) => g.total > 0)
    .sort((a, b) => b.total - a.total || cmp(a.label, b.label));

  return {
    groups: result,
    comboCount: combos.size,
    itemTotalSum: totals.reduce((a, b) => a + b, 0),
    rowCount,
    multiCount,
    unspecified,
  };
}

/** Stable DOM id for a QR group; punctuation is dropped so it is a valid id. */
export const rowElementId = (key: string) => "qr-row-" + key.replace(/[^\p{L}\p{N}]/gu, "");

export function selectedDatesLabel(selected: string[]): string {
  if (selected.length === 0) return "No date selected";
  const sorted = [...selected].sort();
  const latest = sorted[sorted.length - 1];
  if (sorted.length === 1) return formatDay(latest);
  return `${num0(sorted.length)} dates · ${sorted[0]} ~ ${latest}`;
}

export function exportDateText(selected: string[]): string {
  if (selected.length === 0) return "no_date";
  const sorted = [...selected].sort();
  const latest = compactDate(sorted[sorted.length - 1]);
  if (sorted.length === 1) return latest;
  return `${compactDate(sorted[0])}-${latest}_${sorted.length}dates`;
}

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

/** yyyy-MM-dd (ddd), English weekday regardless of the browser's locale. */
export function formatDay(date: string): string {
  const day = new Date(date + "T00:00:00Z").getUTCDay();
  return `${date} (${WEEKDAYS[day]})`;
}

export const displayCell = (value: string) => (value.trim() === "" ? "-" : value);

export function parseDay(text: string | null | undefined): string {
  return text && /^\d{4}-\d{2}-\d{2}$/.test(text.trim()) ? text.trim() : "";
}

// ---------------------------------------------------------------- GAUSS scatter

export const PLOT = { left: 74, right: 890, top: 34, bottom: 500 };
export const CURVE_BASE = 946;
export const CURVE_WIDTH = 150;
export const HISTOGRAM_BINS = 34;
export const BOX_LABEL_PITCH = 12;
export const MAX_PLOTTED_POINTS = 8000;
export const MAX_TITLED_POINTS = 1500;
/** Fraction of a series' slot inside a date column that the jittered strip may fill. */
export const JITTER_FILL = 0.36;
export const MARK_COLOR = "#16202b";
export const VALUE_COLOR = "#1f5fa8";
export const VALUE2_COLOR = "#8a4fb8";

export interface Stats {
  n: number;
  mean: number;
  median: number;
  min: number;
  max: number;
  q1: number;
  q3: number;
  stdDev: number;
}

export const EMPTY_STATS: Stats = { n: 0, mean: 0, median: 0, min: 0, max: 0, q1: 0, q3: 0, stdDev: 0 };

/** Linear interpolation between order statistics (the R type 7 / Excel PERCENTILE rule). */
export function quantile(sorted: number[], q: number): number {
  if (sorted.length === 1) return sorted[0];
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return lo === hi ? sorted[lo] : sorted[lo] + (pos - lo) * (sorted[hi] - sorted[lo]);
}

export function describe(points: GaussPoint[]): Stats {
  const sorted = points.map((p) => p.value).sort((a, b) => a - b);
  if (sorted.length === 0) return EMPTY_STATS;
  const mean = sorted.reduce((a, b) => a + b, 0) / sorted.length;
  const variance = sorted.length < 2 ? 0 : sorted.reduce((a, v) => a + (v - mean) * (v - mean), 0) / (sorted.length - 1);
  return {
    n: sorted.length,
    mean,
    median: quantile(sorted, 0.5),
    min: sorted[0],
    max: sorted[sorted.length - 1],
    q1: quantile(sorted, 0.25),
    q3: quantile(sorted, 0.75),
    stdDev: Math.sqrt(variance),
  };
}

export interface Scale {
  lo: number;
  hi: number;
  ticks: number[];
}

export const UNIT_SCALE: Scale = { lo: 0, hi: 1, ticks: [] };

export function mapY(scale: Scale, value: number): number {
  const span = scale.hi - scale.lo;
  const t = span <= 0 ? 0.5 : (value - scale.lo) / span;
  return PLOT.bottom - t * (PLOT.bottom - PLOT.top);
}

export function niceStep(raw: number): number {
  if (raw <= 0 || !Number.isFinite(raw)) return 1;
  const power = Math.pow(10, Math.floor(Math.log10(raw)));
  const fraction = raw / power;
  const nice = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 2.5 ? 2.5 : fraction <= 5 ? 5 : 10;
  return nice * power;
}

export function buildScale(min: number, max: number): Scale {
  let lo = min;
  let hi = max;
  // A run that never moved would otherwise collapse the axis onto a single pixel.
  if (hi - lo <= 0) {
    const flat = Math.abs(hi) > 0 ? Math.abs(hi) * 0.05 : 1;
    lo -= flat;
    hi += flat;
  } else {
    const pad = (hi - lo) * 0.08;
    lo -= pad;
    hi += pad;
  }
  const step = niceStep((hi - lo) / 6);
  const first = Math.ceil(lo / step) * step;
  const ticks: number[] = [];
  for (let t = first; t <= hi + step * 1e-6 && ticks.length < 14; t += step) ticks.push(t);
  return { lo, hi, ticks };
}

export interface DateColumn {
  index: number;
  date: string;
  label: string;
  count: number;
  left: number;
  center: number;
  showLabel: boolean;
}

/** One column per test date, evenly spaced: a two-day window gets exactly two columns. */
export function buildDateColumns(points: GaussPoint[]): { dates: string[]; columns: DateColumn[]; band: number } {
  const counts = new Map<string, number>();
  for (const p of points) counts.set(p.date, (counts.get(p.date) ?? 0) + 1);
  const dates = [...counts.keys()].sort();
  if (dates.length === 0) return { dates, columns: [], band: PLOT.right - PLOT.left };

  const band = (PLOT.right - PLOT.left) / dates.length;
  // Past roughly 16 columns the MM-dd labels start colliding, so only every nth is drawn.
  const stride = Math.ceil(dates.length / 16);
  const columns = dates.map((date, i): DateColumn => ({
    index: i,
    date,
    label: date.slice(5),
    count: counts.get(date) ?? 0,
    left: PLOT.left + i * band,
    center: PLOT.left + (i + 0.5) * band,
    showLabel: i % stride === 0,
  }));
  return { dates, columns, band };
}

export interface HistogramBar {
  top: number;
  height: number;
  width: number;
}

export function buildHistogram(points: GaussPoint[], scale: Scale): HistogramBar[] {
  const span = scale.hi - scale.lo;
  if (points.length === 0 || span <= 0) return [];
  const binSpan = span / HISTOGRAM_BINS;
  const bins = new Array<number>(HISTOGRAM_BINS).fill(0);
  for (const p of points) {
    const index = Math.trunc((p.value - scale.lo) / binSpan);
    bins[Math.min(Math.max(index, 0), HISTOGRAM_BINS - 1)]++;
  }
  const tallest = Math.max(...bins);
  if (tallest === 0) return [];
  const bars: HistogramBar[] = [];
  bins.forEach((count, i) => {
    if (count === 0) return;
    const binLo = scale.lo + i * binSpan;
    const top = mapY(scale, binLo + binSpan);
    const bottom = mapY(scale, binLo);
    bars.push({ top, height: Math.max(bottom - top, 0.6), width: (count / tallest) * CURVE_WIDTH });
  });
  return bars;
}

const f2 = (value: number) => String(Math.round(value * 100) / 100);

/** Normal curve from mean and σ, drawn vertically: value on the shared Y axis, density to the right. */
export function buildNormalCurve(stats: Stats, scale: Scale): string {
  const span = scale.hi - scale.lo;
  if (stats.n === 0 || stats.stdDev <= 0 || span <= 0) return "";
  const steps = 160;
  const parts: string[] = [];
  for (let i = 0; i <= steps; i++) {
    const v = scale.lo + (span * i) / steps;
    const z = (v - stats.mean) / stats.stdDev;
    parts.push(`${f2(CURVE_BASE + Math.exp(-0.5 * z * z) * CURVE_WIDTH)},${f2(mapY(scale, v))}`);
  }
  return parts.join(" ");
}

export interface BoxLabel {
  text: string;
  lineY: number;
  textY: number;
  fill: string;
}

export const nudged = (label: BoxLabel) => Math.abs(label.textY - 3.4 - label.lineY) > 1.5;

/** Box labels share a narrow column, so they carry at most one decimal. */
export function markNum(value: number): string {
  return Math.abs(value) >= 100
    ? Math.round(value).toLocaleString("en-US")
    : String(Math.round(value * 10) / 10);
}

/**
 * The six numbers beside a box. A tight IQR puts them a few pixels apart, so the column is
 * spread to a readable pitch and each label that moved keeps a leader back to its own line.
 */
export function buildBoxLabels(stats: Stats, accent: string, scale: Scale): BoxLabel[] {
  if (stats.n === 0) return [];
  const source: [string, number, string][] = [
    [`MAX ${markNum(stats.max)}`, stats.max, MARK_COLOR],
    [`Q3 ${markNum(stats.q3)}`, stats.q3, MARK_COLOR],
    [`Q2 ${markNum(stats.median)}`, stats.median, MARK_COLOR],
    [`MEAN ${markNum(stats.mean)}`, stats.mean, accent],
    [`Q1 ${markNum(stats.q1)}`, stats.q1, MARK_COLOR],
    [`MIN ${markNum(stats.min)}`, stats.min, MARK_COLOR],
  ];
  const labels: BoxLabel[] = source
    .map(([text, value, fill]) => ({ text, lineY: mapY(scale, value), textY: mapY(scale, value) + 3.4, fill }))
    .sort((a, b) => a.lineY - b.lineY);

  for (let i = 1; i < labels.length; i++) {
    const floor = labels[i - 1].textY + BOX_LABEL_PITCH;
    if (labels[i].textY < floor) labels[i] = { ...labels[i], textY: floor };
  }

  // Spreading downwards can push the last label under the panel; hang the column off the bottom instead.
  const limit = PLOT.bottom + 13;
  if (labels[labels.length - 1].textY > limit) {
    labels[labels.length - 1] = { ...labels[labels.length - 1], textY: limit };
    for (let i = labels.length - 2; i >= 0; i--) {
      const ceiling = labels[i + 1].textY - BOX_LABEL_PITCH;
      if (labels[i].textY > ceiling) labels[i] = { ...labels[i], textY: ceiling };
    }
  }
  return labels;
}

export interface SigmaMark {
  y: number;
  label: string;
}

export function buildSigmaMarks(stats: Stats, scale: Scale): SigmaMark[] {
  if (stats.n === 0 || stats.stdDev <= 0) return [];
  const marks: SigmaMark[] = [];
  for (const k of [-2, -1, 1, 2]) {
    const v = stats.mean + k * stats.stdDev;
    if (v < scale.lo || v > scale.hi) continue;
    marks.push({ y: mapY(scale, v), label: `${k > 0 ? "+" : "−"}${Math.abs(k)}σ` });
  }
  return marks;
}

/** The markers actually drawn: every point, or an even sample once the strips are solid. */
export function samplePoints<T>(points: T[], budget: number): T[] {
  if (points.length <= budget) return points;
  const stride = Math.ceil(points.length / budget);
  const sampled: T[] = [];
  for (let i = 0; i < points.length; i += stride) sampled.push(points[i]);
  return sampled;
}

export function diamond(cx: number, cy: number): string {
  return `${f2(cx)},${f2(cy - 5)} ${f2(cx + 5)},${f2(cy)} ${f2(cx)},${f2(cy + 5)} ${f2(cx - 5)},${f2(cy)}`;
}

/** Table cells: 1,234.5 / 12.34 / 1.234 depending on magnitude. */
export function num(value: number): string {
  const magnitude = Math.abs(value);
  const digits = magnitude >= 1000 ? 1 : magnitude >= 10 ? 2 : 3;
  return value.toLocaleString("en-US", { minimumFractionDigits: digits, maximumFractionDigits: digits });
}

/** Y axis ticks read as whole gauss values. */
export const axisNum = (value: number) => Math.round(value).toLocaleString("en-US");

export interface SeriesView {
  slot: number;
  lane: number;
  laneCount: number;
  label: string;
  color: string;
  points: GaussPoint[];
  plotted: GaussPoint[];
  stats: Stats;
  ngCount: number;
  ngRatePpm: number | null;
  histogram: HistogramBar[];
  sigmaMarks: SigmaMark[];
  boxLabels: BoxLabel[];
  curve: string;
  boxCenter: number;
  boxHalf: number;
  capHalf: number;
  labelX: number;
  showTitles: boolean;
}

export interface ScatterModel {
  views: SeriesView[];
  scale: Scale;
  dates: string[];
  dateIndex: Map<string, number>;
  columns: DateColumn[];
  band: number;
  totalPoints: number;
  plottedPoints: number;
}

export interface ScatterInput {
  value: GaussPoint[];
  value2: GaussPoint[];
  showValue: boolean;
  showValue2: boolean;
  ngOnly: boolean;
}

/** Re-derives the visible series, the shared axis, the date columns and both panels. */
export function buildScatter(input: ScatterInput): ScatterModel {
  const pick = (points: GaussPoint[]) => (input.ngOnly ? points.filter((p) => p.ng) : points);
  const lanes: { slot: number; label: string; color: string; points: GaussPoint[] }[] = [];
  if (input.showValue) lanes.push({ slot: 0, label: "Value", color: VALUE_COLOR, points: pick(input.value) });
  if (input.showValue2) lanes.push({ slot: 1, label: "Value 2", color: VALUE2_COLOR, points: pick(input.value2) });

  const laneCount = Math.max(lanes.length, 1);
  const total = lanes.reduce((a, l) => a + l.points.length, 0);

  // The two readings share one axis on purpose: side-by-side boxes only compare on the same scale.
  let lo = Number.POSITIVE_INFINITY;
  let hi = Number.NEGATIVE_INFINITY;
  for (const lane of lanes)
    for (const p of lane.points) {
      if (p.value < lo) lo = p.value;
      if (p.value > hi) hi = p.value;
    }
  const scale = total === 0 ? UNIT_SCALE : buildScale(lo, hi);

  const { dates, columns, band } = buildDateColumns(lanes.flatMap((l) => l.points));
  const dateIndex = new Map(dates.map((d, i) => [d, i]));
  const budget = Math.max(Math.floor(MAX_PLOTTED_POINTS / laneCount), 1);
  const single = lanes.length === 1;

  const views = lanes.map((lane, i): SeriesView => {
    const stats = describe(lane.points);
    const raw = lane.slot === 0 ? input.value : input.value2;
    const ng = raw.filter((p) => p.ng).length;
    // Two boxes share the 200px right of the curve: narrower, 100px apart; a single box keeps its old width.
    const boxCenter = single ? 1180 : 1116 + i * 100;
    const boxHalf = single ? 38 : 16;
    return {
      slot: lane.slot,
      lane: i,
      laneCount,
      label: lane.label,
      color: lane.color,
      points: lane.points,
      plotted: samplePoints(lane.points, budget),
      stats,
      ngCount: ng,
      ngRatePpm: raw.length === 0 ? null : (ng / raw.length) * 1_000_000,
      // With one series its histogram sits behind the curve; with two the bars would only muddy each other.
      histogram: single ? buildHistogram(lane.points, scale) : [],
      sigmaMarks: single ? buildSigmaMarks(stats, scale) : [],
      boxLabels: buildBoxLabels(stats, lane.color, scale),
      curve: buildNormalCurve(stats, scale),
      boxCenter,
      boxHalf,
      capHalf: single ? 28 : 12,
      labelX: boxCenter + boxHalf + 12,
      showTitles: lane.points.length <= MAX_TITLED_POINTS,
    };
  });

  return {
    views,
    scale,
    dates,
    dateIndex,
    columns,
    band,
    totalPoints: views.reduce((a, v) => a + v.points.length, 0),
    plottedPoints: views.reduce((a, v) => a + v.plotted.length, 0),
  };
}

/** Each date column is split into one lane per visible reading, so Value and Value 2 stand side by side. */
export function pointX(model: ScatterModel, point: GaussPoint, view: SeriesView): number {
  const index = model.dateIndex.get(point.date) ?? 0;
  const lane = model.band / view.laneCount;
  return PLOT.left + index * model.band + (view.lane + 0.5) * lane + point.jitter * lane * JITTER_FILL;
}

// ---------------------------------------------------------------- windowing

/** Rows to mount for a scroll position: a fixed-height window plus overscan, with spacer heights. */
export function virtualWindow(count: number, itemSize: number, scrollTop: number, viewport: number, overscan = 12) {
  // An even first row keeps the zebra striping stable as the window slides (a spacer row precedes the rows).
  let first = Math.max(Math.floor(scrollTop / itemSize) - overscan, 0);
  first -= first % 2;
  const last = Math.min(Math.ceil((scrollTop + viewport) / itemSize) + overscan, count);
  return { first, last: Math.max(last, first), before: first * itemSize, after: Math.max(count - Math.max(last, first), 0) * itemSize };
}
