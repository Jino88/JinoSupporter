// API of the BMES table pages (Endpoints/BmesTableEndpoints.cs).
import { ApiError, loginUrl } from "./api";

export interface RoutingRow { id: number; modelName: string; processCode: string; processName: string; processType: string }
export interface RoutingEdit { modelName: string; processCode: string; processName: string; processType: string }
export interface ReasonRow { id: number; processName: string; ngName: string; reason: string }
export interface ReasonEdit { processName: string; ngName: string; reason: string }
export interface SyncResult { succeeded: boolean; message: string; rows: number }
export interface Pair { processName: string; ngName: string }
export interface FindPairsResult { apiAdded: number; cacheAdded: number; apiFailure: string | null; cacheFailure: string | null }

export interface WorkerRecord {
  date: string; empNo: string; name: string; departmentNo: string; departmentCode: string; departmentName: string;
  workStatus: string; dayType: string; schedStart: string; schedEnd: string; checkIn: string; checkOut: string;
}
export interface WorkerFetchResult {
  isSuccess: boolean; errorMessage: string; totalCount: number; fetchedAt: string; records: WorkerRecord[];
}

export interface SubGroup { name: string; lineShifts: string[]; subGroups: SubGroup[] }
export interface MidGroup { material: string; subGroups: SubGroup[] }
export interface ModelGroup { name: string; productGroup: string; midGroups: MidGroup[] }
export interface MaterialRow { maktx: string; mtype: string }
export interface ModelGroupView { groups: ModelGroup[]; rows: MaterialRow[]; mtypeMap: Record<string, string> }
export interface PullResult { groupResult: SyncResult; materialResult: SyncResult; view: ModelGroupView }
export interface SaveCounts { groups: number; midGroups: number; subGroups: number; lineShifts: number }

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

async function request(method: string, path: string, body?: unknown, signal?: AbortSignal): Promise<Response> {
  const response = await fetch("/api/bmes/" + path, {
    method,
    credentials: "same-origin",
    signal,
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
  return (await response.json()) as T;
}

async function empty(method: string, path: string, body?: unknown): Promise<void> {
  await request(method, path, body);
}

type StreamLine<T> = { type: "progress"; message: string } | { type: "result"; data: T } | { type: "error"; message: string };

/** POST that answers as NDJSON (progress lines, then the result). Resolves with the result. */
async function stream<T>(path: string, body: unknown, onProgress: (message: string) => void): Promise<T> {
  const response = await request("POST", path, body ?? {});
  const reader = response.body!.getReader();
  const decoder = new TextDecoder();
  let pending = "";
  let result: { data: T } | null = null;
  const handle = (line: string) => {
    if (!line.trim()) return;
    const event = JSON.parse(line) as StreamLine<T>;
    if (event.type === "progress") onProgress(event.message);
    else if (event.type === "error") throw new ApiError(500, event.message);
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
  if (!result) throw new ApiError(500, "The server closed the connection without a result.");
  return (result as { data: T }).data;
}

export const bmesApi = {
  routing: {
    list: () => json<{ rows: RoutingRow[] }>("GET", "routing-table"),
    add: (row: RoutingEdit) => json<{ id: number }>("POST", "routing-table", row),
    update: (id: number, row: RoutingEdit) => empty("PUT", `routing-table/${id}`, row),
    remove: (id: number) => empty("DELETE", `routing-table/${id}`),
    importFile: () => json<{ count: number }>("POST", "routing-table/import", {}),
    refresh: (onProgress: (message: string) => void) => stream<{ saved: number }>("routing-table/refresh", {}, onProgress),
    sync: () => json<SyncResult>("POST", "routing-table/sync", {}),
  },
  reason: {
    list: () => json<{ rows: ReasonRow[] }>("GET", "reason-table"),
    add: (row: ReasonEdit) => json<{ id: number }>("POST", "reason-table", row),
    update: (id: number, row: ReasonEdit) => empty("PUT", `reason-table/${id}`, row),
    remove: (id: number) => empty("DELETE", `reason-table/${id}`),
    importFile: () => json<{ count: number }>("POST", "reason-table/import", {}),
    modelPairs: (model: string) =>
      json<{ pairs: Pair[] }>("GET", "reason-table/model-pairs?model=" + encodeURIComponent(model)),
    findPairs: (onProgress: (message: string) => void) =>
      stream<FindPairsResult>("reason-table/find-pairs", {}, onProgress),
    sync: () => json<SyncResult>("POST", "reason-table/sync", {}),
  },
  worker: {
    fetch: (start: string, end: string, onProgress: (message: string) => void) =>
      stream<WorkerFetchResult>("worker-status/fetch", { start, end }, onProgress),
    exportXlsx: async (start: string, end: string, records: WorkerRecord[]): Promise<Blob> =>
      (await request("POST", "worker-status/export", { start, end, records })).blob(),
  },
  modelGroups: {
    view: () => json<ModelGroupView>("GET", "model-groups"),
    lineShifts: (material: string) =>
      json<{ lineShifts: string[] }>("GET", "model-groups/line-shifts?material=" + encodeURIComponent(material)),
    pull: () => json<PullResult>("POST", "model-groups/pull", {}),
    save: (groups: ModelGroup[]) => json<SaveCounts>("PUT", "model-groups", groups),
  },
};
