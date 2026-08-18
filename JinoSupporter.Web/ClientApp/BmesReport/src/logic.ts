import type {
  BmesReportDocument,
  DailyReasonRow,
  FCostDataset,
  FCostFollowerTabData,
  NullableNumberMap,
  ReportPeriod,
  TabKey,
  ViewerDefaults,
} from "./contract";
import { TAB_KEYS } from "./contract";

export interface ViewerPreferences {
  minimumPpm: number;
  dateColumnLimit: number;
  weekColumnLimit: number;
  monthColumnLimit: number;
}

export const VIEWER_PREFERENCES_KEY = "jinosupporter.bmes-report.viewer.v1";

export const preferencesFromDefaults = (defaults: ViewerDefaults): ViewerPreferences => ({
  minimumPpm: normaliseNonNegative(defaults.minimumPpm, 500),
  dateColumnLimit: normaliseNonNegativeInteger(defaults.dateColumnLimit, 7),
  weekColumnLimit: normaliseNonNegativeInteger(defaults.weekColumnLimit, 4),
  monthColumnLimit: normaliseNonNegativeInteger(defaults.monthColumnLimit, 3),
});

export function parseStoredPreferences(serialized: string | null, defaults: ViewerDefaults): ViewerPreferences {
  const fallback = preferencesFromDefaults(defaults);
  if (!serialized) return fallback;
  try {
    const value = JSON.parse(serialized) as Partial<ViewerPreferences>;
    return {
      minimumPpm: normaliseNonNegative(Number(value.minimumPpm), fallback.minimumPpm),
      dateColumnLimit: normaliseNonNegativeInteger(Number(value.dateColumnLimit), fallback.dateColumnLimit),
      weekColumnLimit: normaliseNonNegativeInteger(Number(value.weekColumnLimit), fallback.weekColumnLimit),
      monthColumnLimit: normaliseNonNegativeInteger(Number(value.monthColumnLimit), fallback.monthColumnLimit),
    };
  } catch {
    return fallback;
  }
}

export function loadStoredPreferences(defaults: ViewerDefaults): ViewerPreferences {
  if (typeof window === "undefined") return preferencesFromDefaults(defaults);
  try {
    return parseStoredPreferences(window.localStorage.getItem(VIEWER_PREFERENCES_KEY), defaults);
  } catch {
    return preferencesFromDefaults(defaults);
  }
}

export function storePreferences(preferences: ViewerPreferences): void {
  if (typeof window === "undefined") return;
  try {
    window.localStorage.setItem(VIEWER_PREFERENCES_KEY, JSON.stringify(preferences));
  } catch {
    // The report remains fully usable when storage is disabled or quota-limited.
  }
}

export function normaliseNonNegative(value: number, fallback: number): number {
  return Number.isFinite(value) && value >= 0 ? value : fallback;
}

export function normaliseNonNegativeInteger(value: number, fallback: number): number {
  const normalised = normaliseNonNegative(value, fallback);
  return Math.floor(normalised);
}

export function selectInitialTab(defaultTab: string): TabKey {
  return TAB_KEYS.includes(defaultTab as TabKey) ? (defaultTab as TabKey) : "daily";
}

export function selectVisiblePeriods(
  periods: ReportPeriod[],
  preferences: Pick<ViewerPreferences, "dateColumnLimit" | "weekColumnLimit" | "monthColumnLimit">,
  allPeriods = false,
): ReportPeriod[] {
  const sorted = [...periods].sort((left, right) => left.sortOrder - right.sortOrder);
  if (allPeriods) return sorted;

  const limitFor = (kind: ReportPeriod["kind"]): number => {
    if (kind === "date") return preferences.dateColumnLimit;
    if (kind === "week") return preferences.weekColumnLimit;
    return preferences.monthColumnLimit;
  };

  return (["date", "week", "month"] as const).flatMap((kind) => {
    const matching = sorted.filter((period) => period.kind === kind);
    const limit = limitFor(kind);
    return limit === 0 ? [] : matching.slice(Math.max(0, matching.length - limit));
  });
}

export function filterDailyReasonRows(
  rows: DailyReasonRow[],
  referencePeriodKey: string | null,
  minimumPpm: number,
): DailyReasonRow[] {
  if (!referencePeriodKey) return rows;
  return rows.filter((row) => {
    if (row.isTotal) return true;
    const ppm = row.ppmByPeriod[referencePeriodKey];
    return ppm !== null && ppm !== undefined && ppm >= minimumPpm;
  });
}

export function sumVisiblePpmByPeriod(rows: DailyReasonRow[], periods: ReportPeriod[]): NullableNumberMap {
  const details = rows.filter((row) => !row.isTotal);
  return Object.fromEntries(
    periods.map((period) => {
      const values = details
        .map((row) => row.ppmByPeriod[period.key])
        .filter((value): value is number => typeof value === "number" && Number.isFinite(value));
      return [period.key, values.length === 0 ? null : values.reduce((sum, value) => sum + value, 0)];
    }),
  );
}

export function resolveFCostDataset(
  report: BmesReportDocument,
  tabData: FCostFollowerTabData,
): FCostDataset | null {
  if (tabData.view.sourceTab !== "fcost") return null;
  return report.tabs.fcost.data?.dataset ?? null;
}

export function isDatasetEmpty(value: unknown): boolean {
  if (value === null || value === undefined) return true;
  if (Array.isArray(value)) return value.length === 0;
  return false;
}
