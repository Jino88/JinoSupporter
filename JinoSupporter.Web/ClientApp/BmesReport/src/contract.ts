export const CONTRACT_ID = "jinosupporter.bmes-report" as const;
export const SUPPORTED_SCHEMA_VERSION = "1.0.0" as const;

export const TAB_KEYS = [
  "daily",
  "weekly",
  "kpi",
  "cause-monthly",
  "fcost",
  "fcost-all",
  "fcost-weekly",
  "fcost-weekly-all",
] as const;

export type TabKey = (typeof TAB_KEYS)[number];
export type ReportState = "complete" | "partial" | "failed";
export type NullableNumberMap = Record<string, number | null>;

export interface ReportIssue {
  code: string;
  message: string;
  source: string;
  retryable: boolean;
}

export interface ReportStatus {
  state: ReportState;
  code: string;
  message: string | null;
  warnings: ReportIssue[];
  errors: ReportIssue[];
}

export interface ReportPeriod {
  key: string;
  kind: "date" | "week" | "month";
  header: string;
  sortOrder: number;
  startDate: string | null;
  endDateExclusive: string | null;
  sourceIndex: number | null;
  sourceCode: string | null;
  sourcePDate: string | null;
}

export interface ReportSelectionMid {
  material: string;
  lineShifts: string[];
}

export interface ReportSelectionGroup {
  id: number;
  name: string;
  midGroups: ReportSelectionMid[];
}

export interface ReportRequest {
  startDate: string;
  endDate: string;
  timeZoneId: string;
  groups: ReportSelectionGroup[];
}

export interface ViewerDefaults {
  defaultTab: string;
  minimumPpm: number;
  dateColumnLimit: number;
  weekColumnLimit: number;
  monthColumnLimit: number;
  fcostCurrency: string;
}

export interface ReportTabEnvelope<T> {
  status: ReportStatus;
  data: T | null;
}

export interface DailySummaryRow {
  rowId: string;
  parentRowId: string | null;
  depth: number;
  level: string;
  groupName: string | null;
  modelName: string | null;
  display: string;
  processType: string | null;
  inputByPeriod: NullableNumberMap;
  ngByPeriod: NullableNumberMap;
  ppmByPeriod: NullableNumberMap;
}

export interface DailyReasonRow {
  rowId: string;
  parentRowId: string | null;
  reason: string;
  rank: number | null;
  isTotal: boolean;
  processType: string | null;
  processName: string | null;
  ngName: string | null;
  ppmByPeriod: NullableNumberMap;
}

export interface DailyModelSection {
  id: string;
  groupName: string;
  modelName: string;
  lineShiftCount: number;
  reasonRows: DailyReasonRow[];
}

export interface DailyTabData {
  periods: ReportPeriod[];
  summaryRows: DailySummaryRow[];
  modelSections: DailyModelSection[];
  referenceDatePeriodKey: string | null;
}

export interface WeeklyTargetRow {
  rowId: string;
  parentRowId: string | null;
  display: string;
  rowKind: string;
  depth: number;
  isLineShift: boolean;
  baselinePpm: number | null;
  targetPpm: number | null;
  achievementPercent: number | null;
  ppmByPeriod: NullableNumberMap;
}

export interface WeeklySummaryRow {
  rowId: string;
  parentRowId: string | null;
  depth: number;
  level: string;
  groupName: string | null;
  modelName: string | null;
  processType: string | null;
  display: string;
  ppmByPeriod: NullableNumberMap;
}

export interface WeeklyTrendSeries {
  id: string;
  label: string;
  groupName: string | null;
  ppmByPeriod: NullableNumberMap;
}

export interface WeeklyDefectRow {
  rank: number;
  lineShift: string | null;
  processType: string;
  processName: string;
  ngName: string;
  ppmByPeriod: NullableNumberMap;
}

export interface WeeklyTabData {
  periods: ReportPeriod[];
  targetRows: WeeklyTargetRow[];
  summaryRows: WeeklySummaryRow[];
  trendSeries: WeeklyTrendSeries[];
  topDefects: WeeklyDefectRow[];
  sortReferencePeriodKey: string | null;
}

export interface KpiLine {
  kind: string;
  label: string;
  unit: MetricUnit;
  annualValue: number | null;
  valuesByPeriod: NullableNumberMap;
}

export interface KpiMetric {
  id: string;
  name: string;
  type: string;
  baselineValue: number | null;
  targetValue: number | null;
  unit: MetricUnit;
  lines: KpiLine[];
}

export interface KpiTabData {
  periods: ReportPeriod[];
  metrics: KpiMetric[];
}

export type MetricUnit = "percent" | "usd" | "ppm" | "none" | string;

export interface CauseRow {
  rowId: string;
  model: string;
  type: string | null;
  process: string | null;
  ngName: string | null;
  number: number | null;
  cause: string | null;
  shareRatio: number | null;
  isSubtotal: boolean;
  ppmByPeriod: NullableNumberMap;
  weightedPpmByPeriod: NullableNumberMap;
}

export interface CauseModelMonthlyRow {
  model: string;
  ppmByPeriod: NullableNumberMap;
}

export interface CauseMonthlyTabData {
  periods: ReportPeriod[];
  rows: CauseRow[];
  modelMonthlyRows: CauseModelMonthlyRow[];
}

export interface FCostView {
  mode: "regular" | "target-defect-rate";
  allPeriods: boolean;
  sourceTab: string | null;
}

export interface FCostTotals {
  inputQtyByPeriod: NullableNumberMap;
  fcostUsdByPeriod: NullableNumberMap;
  ratePercentByPeriod: NullableNumberMap;
}

export interface FCostTrendRow {
  rowId: string;
  parentRowId: string | null;
  depth: number;
  groupName: string;
  modelName: string;
  ngGroupKey: string;
  inputQtyByPeriod: NullableNumberMap;
  fcostUsdByPeriod: NullableNumberMap;
  ngPpmByPeriod: NullableNumberMap;
  fcostSharePercentByPeriod: NullableNumberMap;
}

export interface FCostHierarchyRow {
  rowId: string;
  parentRowId: string | null;
  depth: number;
  groupName: string;
  modelName: string;
  ngGroupKey: string;
  display: string;
  matchedMaterialCount: number;
  inputQtyByPeriod: NullableNumberMap;
  fcostUsdByPeriod: NullableNumberMap;
  sourceRatePercentByPeriod: NullableNumberMap;
}

export interface FCostMaterialRow {
  displayName: string;
  productGroup: string | null;
  modelNo: string | null;
  material: string | null;
  verid: string | null;
  inputQtyByPeriod: NullableNumberMap;
  fcostUsdByPeriod: NullableNumberMap;
  sourceRatePercentByPeriod: NullableNumberMap;
}

export interface FCostRawPrice {
  unitPrice: number | null;
  currency: string | null;
  priceUnit: string | null;
  unitPriceVnd: number | null;
  isMixed: boolean;
}

export interface FCostExchangeRate {
  periodKey: string;
  standardDate: string;
  krwPerUsd: number | null;
  krwPerVnd: number | null;
}

export interface FCostRawMaterialRow {
  groupName: string;
  modelName: string;
  materialCode: string;
  materialName: string;
  fcostVndByPeriod: NullableNumberMap;
  equivalentQtyByPeriod: NullableNumberMap;
  priceByPeriod: Record<string, FCostRawPrice | null>;
  totalFcostVnd: number;
  sourceRows: number;
}

export interface FCostRawBreakdown {
  sourceTable: string;
  nameSource: string;
  warningMessage: string | null;
  periods: ReportPeriod[];
  exchangeRates: FCostExchangeRate[];
  rows: FCostRawMaterialRow[];
}

export interface TargetDefectRateRow {
  modelName: string;
  targetPpm: number | null;
  achievementPercent: number | null;
  actualPpmByPeriod: NullableNumberMap;
}

export interface TargetDefectRate {
  defaultBaselineRatePercent: number;
  defaultTargetRatePercent: number[];
  actionItems: string[];
  rows: TargetDefectRateRow[];
  totalActualPpmByPeriod: NullableNumberMap;
}

export interface FCostDataset {
  periods: ReportPeriod[];
  totals: FCostTotals;
  trendRows: FCostTrendRow[];
  hierarchyRows: FCostHierarchyRow[];
  materials: FCostMaterialRow[];
  unmappedMaterials: FCostMaterialRow[];
  rawBreakdown: FCostRawBreakdown | null;
  targetDefectRate: TargetDefectRate;
}

export interface FCostTabData {
  view: FCostView;
  dataset: FCostDataset;
}

export interface FCostFollowerTabData {
  view: FCostView;
}

export interface BmesReportTabs {
  daily: ReportTabEnvelope<DailyTabData>;
  weekly: ReportTabEnvelope<WeeklyTabData>;
  kpi: ReportTabEnvelope<KpiTabData>;
  "cause-monthly": ReportTabEnvelope<CauseMonthlyTabData>;
  fcost: ReportTabEnvelope<FCostTabData>;
  "fcost-all": ReportTabEnvelope<FCostFollowerTabData>;
  "fcost-weekly": ReportTabEnvelope<FCostFollowerTabData>;
  "fcost-weekly-all": ReportTabEnvelope<FCostFollowerTabData>;
}

export interface BmesReportDocument {
  contractId: string;
  schemaVersion: string;
  calculationVersion: string;
  generatedAtUtc: string;
  request: ReportRequest;
  viewerDefaults: ViewerDefaults;
  status: ReportStatus;
  tabs: BmesReportTabs;
}
