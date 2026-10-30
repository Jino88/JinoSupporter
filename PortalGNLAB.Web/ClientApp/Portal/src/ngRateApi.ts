// API of the NG Rate pages (Endpoints/NgRateEndpoints.cs). Same cookie as the rest of the portal.
import { ApiError, loginUrl } from "./api";

/** PPM per period key; a missing key means 0 (shown as "-"). */
export type Ppm = Record<string, number>;

export interface ChildEntry {
  key: string;
  label: string;
  lsCount: number;
  depth: number;
  /** "MODEL" or "SUB n" */
  level: string;
}

export interface PickerGroup {
  id: number;
  name: string;
  productGroup: string;
  midCount: number;
  lineShiftCount: number;
  /** Model + every named sub (By Model picker). */
  children: ChildEntry[];
  /** Models only (By Section picker). */
  midChildren: ChildEntry[];
}

export interface NgRateSetup {
  credentialsConfigured: boolean;
  loginId: string;
  productGroups: string[];
  groups: PickerGroup[];
}

export interface Period {
  key: string;
  header: string;
}

export interface SummaryRow {
  processType: string;
  isTotal: boolean;
  ppm: Ppm;
}

export interface GroupRow {
  groupName: string;
  ppm: Ppm;
}

export interface ProcessRow {
  rank: number;
  processType: string;
  processName: string;
  ppm: Ppm;
  groups: GroupRow[];
}

export interface NgRow {
  rank: number;
  processType: string;
  processName: string;
  ngName: string;
  ppm: Ppm;
  groups: GroupRow[];
}

export interface ReasonRow {
  reason: string;
  no: number;
  isTotal: boolean;
  processType: string;
  processName: string;
  ngName: string;
  ppm: Ppm;
  groups: GroupRow[];
}

export interface Report {
  dateCols: Period[];
  weekCols: Period[];
  monthCols: Period[];
  summary: SummaryRow[];
  top10Process: ProcessRow[];
  top10Ng: NgRow[];
  reasonRows: ReasonRow[];
  groupNames: string[];
  groupSummary: Record<string, SummaryRow[]>;
  dbPath: string;
  totalRows: number;
  /** HH:mm:ss */
  generatedAt: string;
  modelFilter: string[];
}

export interface LabeledReport {
  label: string;
  report: Report;
}

export interface HierSubRow {
  rowId: string;
  parentId: string | null;
  display: string;
  depth: number;
  hasChildren: boolean;
  isLineShift: boolean;
  ppm: Ppm;
}

export interface HierMid {
  material: string;
  lineShiftCount: number;
  ppm: Ppm;
  rows: HierSubRow[];
}

export interface HierGroup {
  name: string;
  ppm: Ppm;
  mids: HierMid[];
}

export interface ReportResult {
  dbPath: string;
  reports: LabeledReport[];
  hier: HierGroup[];
  exportId: string | null;
}

export interface ProcessTypeRow {
  processType: string;
  ppm: Ppm;
}

export interface Top10Row {
  material: string;
  processType: string;
  processName: string;
  ngName: string;
  datePpm: Ppm;
  weekPpm: Ppm;
  monthPpm: Ppm;
}

export interface ByGroupGroup {
  name: string;
  total: Ppm;
  processTypes: ProcessTypeRow[];
  mids: HierMid[];
  top10: Top10Row[];
}

export interface ByGroupResult {
  dbPath: string;
  dateCols: Period[];
  weekCols: Period[];
  monthCols: Period[];
  totalRows: number;
  groups: ByGroupGroup[];
  top10DateCols: Period[];
  top10WeekCols: Period[];
  top10MonthCols: Period[];
  sortRefKey: string;
  sortRefBucket: string;
  sortRefLabel: string;
  exportId: string | null;
}

export type ReportMode = "by-model" | "by-section";

/** A run that stopped; `panel` = also shown in the setup panel (data fetch failed). */
export class RunError extends ApiError {
  constructor(message: string, public readonly panel: boolean) {
    super(500, message);
  }
}

async function failure(response: Response): Promise<ApiError> {
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    return new ApiError(401, "Signed out");
  }
  let text = `Request failed (HTTP ${response.status})`;
  if (response.status === 403) text = "You do not have permission to open NG Rate.";
  try {
    const data = (await response.json()) as { error?: string };
    if (data?.error) text = data.error;
  } catch {
    // not JSON
  }
  return new ApiError(response.status, text);
}

async function request(method: "GET" | "POST", path: string, body?: unknown): Promise<Response> {
  const response = await fetch("/api/ngrate/" + path, {
    method,
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      ...(method === "GET" ? {} : { "Content-Type": "application/json", "X-Requested-With": "PortalGNLAB" }),
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (!response.ok) throw await failure(response);
  return response;
}

type StreamLine<T> =
  | { type: "progress"; message: string }
  | { type: "result"; data: T }
  | { type: "error"; message: string; panel?: boolean };

/** POST that answers as NDJSON (progress lines, then the result). Resolves with the result. */
async function stream<T>(path: string, body: unknown, onProgress: (message: string) => void): Promise<T> {
  const response = await request("POST", path, body);
  const reader = response.body!.getReader();
  const decoder = new TextDecoder();
  let pending = "";
  let result: { data: T } | null = null;
  const handle = (line: string) => {
    if (!line.trim()) return;
    const event = JSON.parse(line) as StreamLine<T>;
    if (event.type === "progress") onProgress(event.message);
    else if (event.type === "error") throw new RunError(event.message, event.panel === true);
    else result = { data: event.data };
  };
  while (true) {
    const { value, done } = await reader.read();
    pending += decoder.decode(value, { stream: !done });
    const lines = pending.split("\n");
    pending = lines.pop()!;
    lines.forEach(handle);
    if (done) break;
  }
  handle(pending);
  if (!result) throw new RunError("The server closed the connection without a result.", false);
  return (result as { data: T }).data;
}

export const ngRateApi = {
  setup: async () => (await (await request("GET", "setup")).json()) as NgRateSetup,

  report: (
    mode: ReportMode,
    start: string,
    end: string,
    groupIds: number[],
    subEntries: string[],
    onProgress: (message: string) => void,
  ) => stream<ReportResult>("report", { mode, start, end, groupIds, subEntries }, onProgress),

  byGroup: (start: string, end: string, groupNames: string[], onProgress: (message: string) => void) =>
    stream<ByGroupResult>("by-group", { start, end, groupNames }, onProgress),

  /** Workbook of a finished run, with the file name the server picked. */
  async exportExcel(id: string): Promise<{ blob: Blob; fileName: string }> {
    const response = await request("GET", `export/${encodeURIComponent(id)}`);
    const disposition = response.headers.get("Content-Disposition") ?? "";
    const match = /filename\*?=(?:UTF-8'')?"?([^";]+)"?/i.exec(disposition);
    return { blob: await response.blob(), fileName: match ? decodeURIComponent(match[1]) : "NgRate.xlsx" };
  },
};
