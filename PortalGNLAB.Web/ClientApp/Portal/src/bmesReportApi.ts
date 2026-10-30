// API of the BMES report pages (Endpoints/BmesReportEndpoints.cs, /api/bmes-report).
import { ApiError, loginUrl } from "./api";

export interface PickerEntry { key: string; label: string; lsCount: number; depth: number }
export interface PickerGroup { id: number; name: string; productGroup: string; lsCount: number; entries: PickerEntry[] }
export interface Selection { groupIds: number[]; subEntries: string[] }
export interface ReportSetup {
  credentialsConfigured: boolean;
  loginId: string;
  groups: PickerGroup[];
  productGroups: string[];
  activeProductGroup: string;
  selection: Selection;
  stageNames: string[];
}
export type StageStatus = "Pending" | "Running" | "Done" | "Failed";
export interface Stage {
  name: string; status: StageStatus; currentStep: string; percent: number | null;
  stepCurrent: number | null; stepTotal: number | null; phaseName: string | null; elapsedSeconds: number;
}
export interface StagesSnapshot { overall: number; stages: Stage[] }

export interface DailyViewer {
  reactEnabled: boolean; assetsExist: boolean; assetVersion: string;
  scriptUrl: string; cssUrl: string; dataUrl: string; refreshUrl: string;
}

export interface CachedModel { name: string; groups: string }
export interface CatalogModel { productCode: string; productName: string }
export interface ModelCatalog { models: CatalogModel[]; syncedAt: string; workDate: string }
export interface Test5State { cachedModels: CachedModel[]; catalog: ModelCatalog; pdmTotal: number; pdmFetchedAt: string }
export interface BomRow {
  productCode: string; productName: string; materialCode: string; materialName: string;
  bomLevel: number; bomPath: string; bomPathText: string; hasChildren: boolean;
}
export interface BomResult { model: string; rows: BomRow[]; error: string; warning: string; loadSource: string; fetchedAt: string }
export interface BomLoad { results: BomResult[]; cacheLoads: number; serverLoads: number; cachedModels: CachedModel[] }
export interface PdmRow { pdmno: string; pdmvr: string; pdtnoTx: string; pdrtx: string }
export interface CostRow { materialCode: string; unitPrice: number | null; priceCurrency: string; priceSource: string; unitPriceVnd: number | null }
export interface Drawing { success: boolean; error: string; fileName: string; base64: string }

export interface FCostRawStatus { exists: boolean; successCount: number; failedCount: number; totalRows: number; firstDate: string; lastDate: string }
export interface FCostSetup { credentialsConfigured: boolean; loginId: string; rawDbPath: string; rawStatus: FCostRawStatus | null }
export type PeriodKind = "Unknown" | "Day" | "Week" | "Month";
export interface FCostBlock { displayName: string; productGroup: string; material: string; input: number[]; cost: number[] }
export interface FCostSubGroup {
  display: string; groupName: string; midGroupName: string; subGroupKey: string; matchedMaterialCount: number;
  inputByCol: number[]; costByCol: number[]; sourceRateByCol: number[];
}
export interface FCostView {
  queryDate: string; totalRows: number; startDate: string; endDate: string;
  headers: string[]; columns: { header: string; kind: PeriodKind }[];
  total: FCostBlock | null; materials: FCostBlock[]; subGroups: FCostSubGroup[];
  unmapped: { displayName: string; material: string }[];
  rawStatus: FCostRawStatus | null;
}

async function failure(response: Response): Promise<ApiError> {
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    return new ApiError(401, "Signed out");
  }
  let text = `Request failed (HTTP ${response.status})`;
  if (response.status === 403) text = "You do not have permission for this page.";
  try {
    const data = (await response.json()) as { error?: string };
    if (data?.error) text = data.error;
  } catch {
    // not JSON
  }
  return new ApiError(response.status, text);
}

async function request(method: string, path: string, body?: unknown): Promise<Response> {
  const response = await fetch("/api/bmes-report/" + path, {
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

async function json<T>(method: string, path: string, body?: unknown): Promise<T> {
  const response = await request(method, path, body);
  if (response.status === 204) return undefined as T;
  return (await response.json()) as T;
}

/** One line of an NDJSON answer other than the final result. */
export type StreamEvent =
  | { type: "progress"; message: string }
  | { type: "log"; message: string }
  | ({ type: "stages" } & StagesSnapshot);

type StreamLine<T> = StreamEvent | { type: "result"; data: T } | { type: "error"; message: string };

/** Splits an NDJSON body into lines as they arrive; `feed` returns the complete lines. */
export function lineSplitter() {
  let pending = "";
  return {
    feed(chunk: string): string[] {
      pending += chunk;
      const lines = pending.split("\n");
      pending = lines.pop()!;
      return lines.filter((l) => l.trim());
    },
    rest(): string[] {
      const last = pending;
      pending = "";
      return last.trim() ? [last] : [];
    },
  };
}

/** POST that answers as NDJSON (progress/log/stages lines, then the result). Resolves with the result. */
async function stream<T>(path: string, body: unknown, onEvent: (event: StreamEvent) => void): Promise<T> {
  const response = await request("POST", path, body ?? {});
  const reader = response.body!.getReader();
  const decoder = new TextDecoder();
  const splitter = lineSplitter();
  let result: { data: T } | null = null;
  const handle = (line: string) => {
    const event = JSON.parse(line) as StreamLine<T>;
    if (event.type === "result") result = { data: event.data };
    else if (event.type === "error") throw new ApiError(500, event.message);
    else onEvent(event);
  };
  while (true) {
    const { value, done } = await reader.read();
    splitter.feed(decoder.decode(value, { stream: !done })).forEach(handle);
    if (done) break;
  }
  splitter.rest().forEach(handle);
  if (!result) throw new ApiError(500, "The server closed the connection without a result.");
  return (result as { data: T }).data;
}

/** Progress text of a stream event, for pages that only show one status line. */
const progressOnly = (onProgress: (message: string) => void) => (event: StreamEvent) => {
  if (event.type === "progress") onProgress(event.message);
};

export const bmesReportApi = {
  daily: {
    viewer: () => json<DailyViewer>("GET", "daily/viewer"),
    viewerFailed: (message: string) => json<void>("POST", "daily/viewer-failed", { message }),
  },
  report: {
    setup: () => json<ReportSetup>("GET", "report/setup"),
    saveSelection: (selection: Selection) => json<Selection>("PUT", "report/selection", selection),
    generate: (start: string, end: string, selection: Selection, onEvent: (event: StreamEvent) => void) =>
      stream<{ token: string }>("report/generate", { start, end, ...selection }, onEvent),
  },
  test5: {
    state: (plant: string) => json<Test5State>("GET", "test5/state?plant=" + encodeURIComponent(plant)),
    catalog: (plant: string) => json<ModelCatalog>("GET", "test5/catalog?plant=" + encodeURIComponent(plant)),
    syncCatalog: (plant: string, workDate: string) => json<ModelCatalog>("POST", "test5/catalog/sync", { plant, workDate }),
    loadBom: (
      body: { models: string[]; plant: string; workDate: string; maxDepth: number; fromServer: boolean },
      onProgress: (message: string) => void,
    ) => stream<BomLoad>("test5/bom", body, progressOnly(onProgress)),
    pdmMatch: (partNames: string[], modelNames: string[]) =>
      json<{ rows: PdmRow[] }>("POST", "test5/pdm/match", { partNames, modelNames }),
    fetchPdm: (onProgress: (message: string) => void) =>
      stream<{ saved: number; total: number; fetchedAt: string }>("test5/pdm/fetch", {}, progressOnly(onProgress)),
    drawing: (pdmno: string, pdmvr: string, onProgress: (message: string) => void) =>
      stream<Drawing>("test5/pdm/drawing", { pdmno, pdmvr }, progressOnly(onProgress)),
    costs: (codes: string[]) => json<{ rows: CostRow[]; error: string }>("POST", "test5/costs", { codes }),
    activity: (message: string) => json<void>("POST", "test5/activity", { message }),
  },
  fcost: {
    setup: () => json<FCostSetup>("GET", "f-cost/setup"),
    generate: (start: string, end: string, onProgress: (message: string) => void) =>
      stream<FCostView>("f-cost/generate", { start, end }, progressOnly(onProgress)),
    activity: (message: string) => json<void>("POST", "f-cost/activity", { message }),
  },
};
