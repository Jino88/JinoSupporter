// Production Input (Endpoints/ProductionInputEndpoints.cs). The server keeps setup, day data,
// day summary and plan as JSON verbatim; their shape is defined here only.

/** Fixed flow blocks of the line; see SEGMENTS in logic.ts for their start rules. */
export type Segment = "A" | "B" | "C" | "D" | "E" | "F";
/** P = process NG, F = function NG (the two NG rate totals). */
export type NgGroup = "P" | "F";
/** Today's value of a plan stage that comes from the daily input. */
export type StageSource = "sub" | "main" | "func" | "vis" | "oqc";
export type Slot = "day" | "am" | "pm";

export interface NgItem {
  id: string;
  name: string;
}

export interface ProcessMaterial {
  material: string;
  /** Pieces used per 1 unit going into the process. */
  qtyPerUnit: number;
}

export interface Process {
  id: string;
  name: string;
  localName: string;
  segment: Segment;
  group: NgGroup;
  ngItems: NgItem[];
  materials: ProcessMaterial[];
}

export interface PlanStage {
  id: string;
  name: string;
  /** null = the actual is typed in the plan table. */
  source: StageSource | null;
}

export interface ModelSetup {
  processes: Process[];
  materials: string[];
  stages: PlanStage[];
}

export interface ModelSummary {
  id: string;
  name: string;
  build: string;
  processCount: number;
  materialCount: number;
  revision: number;
  updatedAt: string;
  updatedBy: string;
}

export interface Model {
  id: string;
  name: string;
  build: string;
  setup: ModelSetup;
  revision: number;
  updatedAt: string;
  updatedBy: string;
}

/** [before rework, after rework] NG count. */
export type Pair = [number, number];
/** Process id → NG item id → pair. */
export type NgCounts = Record<string, Record<string, Pair>>;

export interface MaterialDay {
  /** Typed opening stock; null = previous day's closing. */
  opening: number | null;
  in: number;
  source: string;
  /** Typed out; null = Σ process input × qty per unit. */
  out: number | null;
  counted: number | null;
}

export interface DayData {
  /** Typed start input of the segments that start from a count (A, C). */
  startSet: Partial<Record<Segment, number>>;
  /** First function process id → input tested in 08–12h. */
  amShare: Record<string, number>;
  /** Process id → typed input instead of the automatic one. */
  override: Record<string, number>;
  /** NG of the non-function segments (one count per day). */
  ng: NgCounts;
  /** NG of segment E per time slot; the day value is am + pm. */
  am: NgCounts;
  pm: NgCounts;
  sendOqc: number;
  materials: Record<string, MaterialDay>;
  /** WIP stage key → floor count. */
  wipCount: Record<string, number | null>;
}

/** Saved with each day so later days, the plan and the reports do not recompute old days. */
export interface DaySummary {
  sub: number | null;
  main: number | null;
  func: number | null;
  vis: number | null;
  oqc: number | null;
  /** Output of the last process of each segment. */
  segOut: Partial<Record<Segment, number | null>>;
  /** Start inputs of the "prev" segments used that day. */
  prevOut: Partial<Record<Segment, number>>;
  /** Material closing stock (opening + in − out). */
  closing: Record<string, number>;
  /** Stock carried to the next day: the counted stock when entered, otherwise the closing. */
  stock?: Record<string, number>;
}

export type DayStatus = "Draft" | "Submitted";

export interface Day {
  modelId: string;
  date: string;
  status: DayStatus;
  data: DayData;
  summary: DaySummary;
  revision: number;
  enteredBy: string;
  updatedAt: string;
}

export interface PlanData {
  /** Stage id → date → planned quantity. */
  values: Record<string, Record<string, number>>;
  /** Stage id → date → typed actual (stages without a source). */
  actuals: Record<string, Record<string, number>>;
  /** Date → material arrival note. */
  events: Record<string, string>;
}

export interface Plan {
  data: PlanData | null;
  revision: number;
  updatedAt: string;
  updatedBy: string;
}

/** A sheet for POST /export and the on-screen report tables. Merges are 0-based. */
export interface SheetMerge {
  row: number;
  col: number;
  rowSpan: number;
  colSpan: number;
}

export type Cell = string | number | null;

export interface Sheet {
  name: string;
  rows: Cell[][];
  merges: SheetMerge[];
  headerRows: number;
  colWidths?: number[];
}

export const emptyDayData = (): DayData => ({
  startSet: {},
  amShare: {},
  override: {},
  ng: {},
  am: {},
  pm: {},
  sendOqc: 0,
  materials: {},
  wipCount: {},
});

export const emptyPlanData = (): PlanData => ({ values: {}, actuals: {}, events: {} });
