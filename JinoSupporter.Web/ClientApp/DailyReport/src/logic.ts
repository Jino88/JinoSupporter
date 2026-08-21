import type { DailyReportModelRow, DailyReportSnapshot } from "./contract";

export type Granularity = "week" | "day";
export type SortKey = "delta" | "latest" | "qty" | "name";
export type GroupFilter = "ALL" | "SPK" | "MODULE" | "UNIT" | "TWS";
export const TREND_DIRECTION = "latest-left" as const;

export interface GranularityData {
  noun: "week" | "day";
  stepUnit: "w" | "d";
  headers: string[];
  recentHeader: string;
  previousHeader: string;
  overallRecentPpm: number;
  overallPreviousPpm: number;
  worsened: DailyReportModelRow[];
  improved: DailyReportModelRow[];
}

export function selectGranularity(snapshot: DailyReportSnapshot, granularity: Granularity): GranularityData {
  if (granularity === "day") {
    return {
      noun: "day",
      stepUnit: "d",
      headers: snapshot.dayHeaders,
      recentHeader: snapshot.recentDayHeader,
      previousHeader: snapshot.previousDayHeader,
      overallRecentPpm: snapshot.overallRecentDayPpm,
      overallPreviousPpm: snapshot.overallPreviousDayPpm,
      worsened: snapshot.worsenedDaily,
      improved: snapshot.improvedDaily,
    };
  }
  return {
    noun: "week",
    stepUnit: "w",
    headers: snapshot.weekHeaders,
    recentHeader: snapshot.recentWeekHeader,
    previousHeader: snapshot.previousWeekHeader,
    overallRecentPpm: snapshot.overallRecentPpm,
    overallPreviousPpm: snapshot.overallPreviousPpm,
    worsened: snapshot.worsened,
    improved: snapshot.improved,
  };
}

export interface RowFilter {
  group: GroupFilter;
  query: string;
  sort: SortKey;
}

export function filterAndSortRows(rows: DailyReportModelRow[], filter: RowFilter): DailyReportModelRow[] {
  const query = filter.query.trim().toLocaleLowerCase();
  const visible = rows.filter((row) => {
    if (filter.group !== "ALL" && row.groupName.trim().toUpperCase() !== filter.group) return false;
    if (!query) return true;
    return `${row.modelName} ${row.parentModelName} ${row.groupName}`.toLocaleLowerCase().includes(query);
  });

  return visible.sort((a, b) => {
    if (filter.sort === "latest") return b.recentPpm - a.recentPpm;
    if (filter.sort === "qty") return b.recentNgQty - a.recentNgQty;
    if (filter.sort === "name") return a.modelName.localeCompare(b.modelName);
    return Math.abs(b.deltaPpm) - Math.abs(a.deltaPpm);
  });
}

export function baselineGapLabel(stepsBack: number, unit: "w" | "d"): string | null {
  return stepsBack > 1 ? `−${stepsBack}${unit}` : null;
}

export interface SparkPoint {
  x: number;
  y: number;
  value: number;
  sourceIndex: number;
}

export interface SparkGeometry {
  width: number;
  height: number;
  zeroY: number;
  baselineY: number | null;
  segments: SparkPoint[][];
  latestPoint: SparkPoint;
}

export function buildSparkGeometry(
  newestFirstValues: number[],
  baselinePpm: number,
  width = 112,
  height = 30,
  padding = 3.5,
): SparkGeometry | null {
  if (newestFirstValues.length < 2) return null;

  // The service contract is newest-first and the user requires that order on screen:
  // latest at x=padding on the left, progressively older periods to the right.
  const values = newestFirstValues;
  const positive = values.filter((value) => value > 0);
  if (positive.length < 2 || values[0] <= 0) return null;

  const high = Math.max(...positive, baselinePpm);
  const stepX = (width - padding * 2) / (values.length - 1);
  const y = (value: number) => height - padding - value / high * (height - padding * 2);

  const segments: SparkPoint[][] = [];
  let current: SparkPoint[] = [];
  values.forEach((value, sourceIndex) => {
    if (value > 0) {
      current.push({ x: padding + stepX * sourceIndex, y: y(value), value, sourceIndex });
      return;
    }
    if (current.length > 0) segments.push(current);
    current = [];
  });
  if (current.length > 0) segments.push(current);

  return {
    width,
    height,
    zeroY: y(0),
    baselineY: baselinePpm > 0 ? y(baselinePpm) : null,
    segments,
    latestPoint: segments[0][0],
  };
}
