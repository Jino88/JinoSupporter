import { validateDailyReportEnvelope, type DailyReportEnvelope } from "./contract";

export type DailyReportLoadErrorKind = "unauthorized" | "forbidden" | "invalid-data" | "network";

export class DailyReportLoadError extends Error {
  constructor(
    public readonly kind: DailyReportLoadErrorKind,
    message: string,
    public readonly status?: number,
  ) {
    super(message);
    this.name = "DailyReportLoadError";
  }
}

async function requestEnvelope(
  url: string,
  method: "GET" | "POST",
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch,
): Promise<DailyReportEnvelope> {
  let response: Response;
  try {
    response = await fetcher(url, {
      method,
      credentials: "same-origin",
      cache: "no-store",
      headers: {
        Accept: "application/json",
        "X-Requested-With": "PortalGNLAB.DailyReport",
      },
      signal,
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") throw error;
    throw new DailyReportLoadError("network", "DAILY REPORT server is unavailable.");
  }

  if (!response.ok) {
    if (response.status === 401) {
      throw new DailyReportLoadError("unauthorized", "Your login session has expired.", 401);
    }
    if (response.status === 403) {
      throw new DailyReportLoadError("forbidden", "You do not have permission to view DAILY REPORT.", 403);
    }
    throw new DailyReportLoadError("network", `Could not load DAILY REPORT. (HTTP ${response.status})`, response.status);
  }

  let payload: unknown;
  try {
    payload = await response.json();
  } catch {
    throw new DailyReportLoadError("invalid-data", "DAILY REPORT response is not valid JSON.");
  }
  return validateDailyReportEnvelope(payload);
}

export function fetchDailyReport(
  dataUrl: string,
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch,
): Promise<DailyReportEnvelope> {
  return requestEnvelope(dataUrl, "GET", signal, fetcher);
}

export function refreshDailyReport(
  refreshUrl: string,
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch,
): Promise<DailyReportEnvelope> {
  return requestEnvelope(refreshUrl, "POST", signal, fetcher);
}
