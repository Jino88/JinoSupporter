import { ApiError, loginUrl } from "./api";

// JSON API of the QR pages (Endpoints/QrEndpoints.cs). Same cookie as the rest of the portal.

export type Verdict = "pass" | "fail" | "unknown";

export interface QrPrefs {
  configured: boolean;
  connection: { server: string; port: number; database: string; userId: string };
  defaultMaxRows: number;
  /** Last picked row cap on a data page; 0 = never picked, -1 = auto. */
  maxRows: number;
  /** Last picked QR cap on SUMMARY / the scatter; 0 = never picked. */
  maxQr: number;
  modelCodes: { code: string; name: string }[];
}

export interface QrItem {
  label: string;
  value: string;
  verdict: Verdict;
}

export interface QrCell {
  station: string;
  verdict: Verdict;
  result: string;
  /** yyyy-MM-dd HH:mm:ss */
  time: string | null;
  items: QrItem[];
}

export interface QrSummaryRow {
  productId: string;
  round: number;
  latestTime: string | null;
  finalPass: boolean;
  finalNote: string;
  hasAnyFail: boolean;
  complete: boolean;
  airleak: QrCell | null;
  bako: QrCell | null;
  gauss: QrCell | null;
  older: QrSummaryRow[];
}

export interface QrSummary {
  from: string;
  to: string;
  maxRows: number;
  totalProducts: number;
  hitMaxRows: boolean;
  maxDate: string | null;
  fetchedAt: string;
  warning: string;
  rows: QrSummaryRow[];
}

export interface QrColumn {
  name: string;
  dataType: string;
  numeric: boolean;
}

export interface QrDataRow {
  values: string[];
  /** yyyy-MM-dd, or null when the row has no test time. */
  date: string | null;
}

export interface QrData {
  table: { key: string; station: string; menuLabel: string; displayTable: string; passValues: string[]; failValues: string[] };
  fetchedAt: string;
  maxRows: number;
  totalRows: number | null;
  validTotalRows: number | null;
  excludedRows: number | null;
  sortColumnName: string;
  selectedDates: string[];
  warning: string;
  hitMaxRows: boolean;
  summary: { inputCount: number; ngCount: number; ngRatePpm: number | null } | null;
  dates: { date: string; rowCount: number }[];
  columns: QrColumn[];
  /** Display order: column index, or -1 for the derived test-date column. */
  layout: number[];
  rows: QrDataRow[];
}

export interface GaussPoint {
  productId: string;
  /** yyyy-MM-dd HH:mm:ss */
  time: string;
  date: string;
  value: number;
  ng: boolean;
  jitter: number;
}

export interface GaussScatter {
  from: string;
  to: string;
  maxRows: number;
  maxDate: string | null;
  fetchedAt: string;
  warning: string;
  productIds: string[];
  value: GaussPoint[];
  value2: GaussPoint[];
  rawValueCount: number;
  rawValue2Count: number;
  ngValueCount: number;
  ngValue2Count: number;
}

async function check(response: Response): Promise<Response> {
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    throw new ApiError(401, "Signed out");
  }
  if (response.status === 403) throw new ApiError(403, "You do not have permission to open this QR page.");
  if (!response.ok) {
    const body = await response.json().catch(() => null);
    throw new ApiError(response.status, body?.error ?? `Request failed (HTTP ${response.status})`);
  }
  return response;
}

async function getQr<T>(path: string, signal?: AbortSignal): Promise<T> {
  const response = await check(
    await fetch("/api/qr/" + path, { credentials: "same-origin", signal, headers: { Accept: "application/json" } }),
  );
  return (await response.json()) as T;
}

export function query(params: Record<string, string | number | undefined>): string {
  const parts = Object.entries(params)
    .filter(([, value]) => value !== undefined && value !== "")
    .map(([key, value]) => `${key}=${encodeURIComponent(String(value))}`);
  return parts.length ? "?" + parts.join("&") : "";
}

export const qrApi = {
  prefs: () => getQr<QrPrefs>("prefs"),

  async savePrefs(change: { maxRows?: number; maxQr?: number }): Promise<void> {
    await check(
      await fetch("/api/qr/prefs", {
        method: "PUT",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json", "X-Requested-With": "PortalGNLAB" },
        body: JSON.stringify(change),
      }),
    );
  },

  summary: (from: string, to: string, max: number, signal?: AbortSignal) =>
    getQr<QrSummary>("summary" + query({ from, to, max }), signal),

  data: (key: string, dates: string[], maxRows: number, signal?: AbortSignal) =>
    getQr<QrData>(`data/${encodeURIComponent(key)}` + query({ dates: dates.join(","), maxRows }), signal),

  scatter: (from: string, to: string, max: number, signal?: AbortSignal) =>
    getQr<GaussScatter>("gauss-scatter" + query({ from, to, max }), signal),

  /** Workbook from an unlimited server-side read; `truncated` is set when it stopped at the export ceiling. */
  async exportSummary(from: string, to: string): Promise<{ blob: Blob; truncated: boolean }> {
    return exportFile("summary/export" + query({ from, to }));
  },

  async exportData(key: string, dates: string[]): Promise<{ blob: Blob; truncated: boolean }> {
    return exportFile(`data/${encodeURIComponent(key)}/export` + query({ dates: dates.join(",") }));
  },
};

async function exportFile(path: string): Promise<{ blob: Blob; truncated: boolean }> {
  const response = await check(await fetch("/api/qr/" + path, { credentials: "same-origin" }));
  return { blob: await response.blob(), truncated: response.headers.get("X-Qr-Truncated") === "1" };
}

/** Saves a blob through a temporary link. */
export function saveBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** window.open that reports whether the browser let it through (pop-up blockers return null). */
export function openWindow(url: string, features: string, name?: string): boolean {
  const opened = window.open(url, name ?? "_blank", features);
  if (!opened) return false;
  opened.focus();
  return true;
}
