export const SUPPORTED_SCHEMA_VERSION = 9 as const;

export interface DailyReportCause {
  ngName: string;
  processName: string;
  processType: string;
  ppm: number;
  previousPpm: number;
  deltaPpm: number;
  deltaPercent: number;
  isNew: boolean;
}

export interface DailyReportModelRow {
  modelName: string;
  groupName: string;
  parentModelName: string;
  recentPpm: number;
  previousPpm: number;
  deltaPpm: number;
  deltaPercent: number;
  trendPpm: number[];
  topCauses: DailyReportCause[];
  recentNgQty: number;
  baselineHeader: string;
  baselineStepsBack: number;
}

export interface DailyReportSnapshot {
  schemaVersion: number;
  generatedAt: string;
  startDate: string;
  endDate: string;
  weekHeaders: string[];
  recentWeekHeader: string;
  previousWeekHeader: string;
  overallRecentPpm: number;
  overallPreviousPpm: number;
  dayHeaders: string[];
  recentDayHeader: string;
  previousDayHeader: string;
  overallRecentDayPpm: number;
  overallPreviousDayPpm: number;
  modelCount: number;
  worsened: DailyReportModelRow[];
  improved: DailyReportModelRow[];
  worsenedDaily: DailyReportModelRow[];
  improvedDaily: DailyReportModelRow[];
  error: string | null;
}

export interface DailyReportEnvelope {
  snapshot: DailyReportSnapshot | null;
  isRefreshing: boolean;
  logs: string[];
}

export class DailyReportContractError extends Error {
  constructor(
    public readonly kind: "unsupported-schema" | "invalid-data",
    message: string,
  ) {
    super(message);
    this.name = "DailyReportContractError";
  }
}

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNumberArray = (value: unknown): value is number[] =>
  Array.isArray(value) && value.every((item) => typeof item === "number" && Number.isFinite(item));

const isStringArray = (value: unknown): value is string[] =>
  Array.isArray(value) && value.every((item) => typeof item === "string");

function isCause(value: unknown): value is DailyReportCause {
  if (!isObject(value)) return false;
  return ["ngName", "processName", "processType"].every((key) => typeof value[key] === "string")
    && ["ppm", "previousPpm", "deltaPpm", "deltaPercent"].every(
      (key) => typeof value[key] === "number" && Number.isFinite(value[key]),
    )
    && typeof value.isNew === "boolean";
}

function isModelRow(value: unknown): value is DailyReportModelRow {
  if (!isObject(value)) return false;
  return ["modelName", "groupName", "parentModelName", "baselineHeader"].every(
    (key) => typeof value[key] === "string",
  )
    && ["recentPpm", "previousPpm", "deltaPpm", "deltaPercent", "recentNgQty", "baselineStepsBack"].every(
      (key) => typeof value[key] === "number" && Number.isFinite(value[key]),
    )
    && isNumberArray(value.trendPpm)
    && Array.isArray(value.topCauses)
    && value.topCauses.every(isCause);
}

function validateSnapshot(value: unknown): DailyReportSnapshot {
  if (!isObject(value)) {
    throw new DailyReportContractError("invalid-data", "DAILY REPORT snapshot is not an object.");
  }
  if (value.schemaVersion !== SUPPORTED_SCHEMA_VERSION) {
    throw new DailyReportContractError(
      "unsupported-schema",
      `Unsupported DAILY REPORT schema: server ${String(value.schemaVersion)}, viewer ${SUPPORTED_SCHEMA_VERSION}.`,
    );
  }

  const strings = [
    "generatedAt", "startDate", "endDate", "recentWeekHeader", "previousWeekHeader",
    "recentDayHeader", "previousDayHeader",
  ];
  const numbers = [
    "overallRecentPpm", "overallPreviousPpm", "overallRecentDayPpm",
    "overallPreviousDayPpm", "modelCount",
  ];
  const rowLists = ["worsened", "improved", "worsenedDaily", "improvedDaily"];
  if (!strings.every((key) => typeof value[key] === "string")
      || !numbers.every((key) => typeof value[key] === "number" && Number.isFinite(value[key]))
      || !isStringArray(value.weekHeaders)
      || !isStringArray(value.dayHeaders)
      || !rowLists.every((key) => Array.isArray(value[key]) && value[key].every(isModelRow))
      || !(value.error === null || typeof value.error === "string")) {
    throw new DailyReportContractError("invalid-data", "DAILY REPORT snapshot fields are invalid.");
  }

  return value as unknown as DailyReportSnapshot;
}

export function validateDailyReportEnvelope(value: unknown): DailyReportEnvelope {
  if (!isObject(value)
      || typeof value.isRefreshing !== "boolean"
      || !isStringArray(value.logs)
      || !(value.snapshot === null || isObject(value.snapshot))) {
    throw new DailyReportContractError("invalid-data", "DAILY REPORT response fields are invalid.");
  }
  return {
    snapshot: value.snapshot === null ? null : validateSnapshot(value.snapshot),
    isRefreshing: value.isRefreshing,
    logs: value.logs,
  };
}
