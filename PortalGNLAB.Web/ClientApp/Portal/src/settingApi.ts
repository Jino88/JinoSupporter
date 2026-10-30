// Setting API (Endpoints/SettingEndpoints.cs) for pages/Setting.tsx.
import { ApiError, getJson, loginUrl } from "./api";

export interface PathsView {
  dbRootDirectory: string;
  defaultDbRootDirectory: string;
  ngRateFolderName: string;
  fCostFolderName: string;
  workerStatusFolderName: string;
  settingsDbPath: string;
}

export interface Credentials {
  loginId: string;
  password: string;
  apiFacility: string;
  apiLoginId: string;
  apiPassword: string;
}

export interface HistoryView {
  isRunning: boolean;
  status: string;
  log: string[];
  lastHistoryDay: string;
  liveWindowDays: number;
}

export interface BmesMaterial {
  matnr: string;
  maktx: string;
  mtype: string;
  btype: string;
  modNameB: string;
  lotQt: string;
  bunch: string;
  ernam: string;
  erdat: string;
  fetchedAt: string;
}

export interface MaterialView {
  latest: BmesMaterial | null;
  count: number;
}

export interface CodexView {
  apiKey: string;
  translateModel: string;
  translateEffort: string;
  reportModel: string;
  reportEffort: string;
  reasoningEfforts: string[];
  lunaModel: string;
  defaultReportModel: string;
  defaultReportEffort: string;
}

export interface AiProviders {
  claudeCliEnabled: boolean;
  claudeApiEnabled: boolean;
  codexCliEnabled: boolean;
  codexApiEnabled: boolean;
}

export interface QrDb {
  server: string;
  port: number;
  database: string;
  userId: string;
  password: string;
  timeoutSeconds: number;
  defaultMaxRows: number;
  encrypt: boolean;
  trustServerCertificate: boolean;
}

export interface QrModelCode {
  code: string;
  name: string;
}

export interface QrModelsView {
  codes: QrModelCode[];
  path: string;
}

export interface SettingView {
  isAdmin: boolean;
  paths: PathsView;
  credentials: Credentials;
  history: HistoryView;
  material: MaterialView;
  codex: CodexView | null;
  aiProviders: AiProviders | null;
  qrDb: QrDb | null;
  qrModels: QrModelsView | null;
}

export interface AppPathResult {
  paths: PathsView;
  message: string;
  requiresRestart: boolean;
}

export interface ServerTableResult {
  succeeded: boolean;
  message: string;
  paths: PathsView;
  requiresRestart: boolean;
}

export interface MaterialFetchEvent {
  log?: string;
  done?: boolean;
  saved?: number;
  error?: string;
  material?: MaterialView;
}

const BASE = "/api/portal/setting";

async function post<T>(path: string, body?: unknown, onLine?: (line: unknown) => void): Promise<T> {
  const response = await fetch(BASE + path, {
    method: "POST",
    credentials: "same-origin",
    headers: { Accept: "application/json", "Content-Type": "application/json", "X-Requested-With": "PortalGNLAB" },
    body: JSON.stringify(body ?? {}),
  });
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname));
    throw new ApiError(401, "Signed out");
  }
  if (!response.ok) {
    const error = (await response.json().catch(() => null)) as { error?: string } | null;
    throw new ApiError(response.status, error?.error ?? `Request failed (HTTP ${response.status})`);
  }
  if (onLine && response.body) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let pending = "";
    let last: unknown = null;
    const emit = (line: string) => {
      if (!line.trim()) return;
      last = JSON.parse(line);
      onLine(last);
    };
    for (;;) {
      const { value, done } = await reader.read();
      pending += decoder.decode(value, { stream: !done });
      const lines = pending.split("\n");
      pending = lines.pop() ?? "";
      lines.forEach(emit);
      if (done) break;
    }
    emit(pending);
    return last as T;
  }
  return (await response.json()) as T;
}

export const settingApi = {
  load: () => getJson<SettingView>(BASE),
  codexModels: () => getJson<{ models: string[]; error: string }>(BASE + "/codex-models"),
  saveCodex: (body: Pick<CodexView, "apiKey" | "translateModel" | "translateEffort" | "reportModel" | "reportEffort">) =>
    post<{ codex: CodexView; aiProviders: AiProviders }>("/codex", body),
  saveAiProviders: (body: { codexCliEnabled: boolean; codexApiEnabled: boolean }) => post<AiProviders>("/ai-providers", body),
  saveQrDb: (body: QrDb) => post<{ qrDb: QrDb; message: string }>("/qr-db", body),
  testQrDb: (body: QrDb) => post<{ qrDb: QrDb; message: string; error: boolean }>("/qr-db/test", body),
  saveQrModels: (codes: QrModelCode[]) => post<QrModelsView>("/qr-model-codes", codes),
  saveAppPath: (dbRootDirectory: string) => post<AppPathResult>("/app-path", { dbRootDirectory }),
  saveCredentials: (body: { loginId: string; password: string }) => post<{ message: string }>("/credentials", body),
  saveApiCredentials: (body: { apiFacility: string; apiLoginId: string; apiPassword: string }) =>
    post<{ message: string }>("/api-credentials", body),
  pullServerTable: (table: "routing" | "reason" | "all", dbRootDirectory: string) =>
    post<ServerTableResult>(`/server-tables/${table}`, { dbRootDirectory }),
  history: () => getJson<HistoryView>(BASE + "/history"),
  startHistory: (from: string) => post<HistoryView>("/history/start", { from }),
  cancelHistory: () => post<HistoryView>("/history/cancel"),
  fetchMaterials: (onEvent: (event: MaterialFetchEvent) => void) =>
    post<MaterialFetchEvent>("/materials/fetch", {}, (line) => onEvent(line as MaterialFetchEvent)),
};
