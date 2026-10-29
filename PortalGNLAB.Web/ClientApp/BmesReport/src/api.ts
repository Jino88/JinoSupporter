import {
  CONTRACT_ID,
  SUPPORTED_SCHEMA_VERSION,
  TAB_KEYS,
  type BmesReportDocument,
} from "./contract";

export type ReportLoadErrorKind =
  | "expired"
  | "unauthorized"
  | "forbidden"
  | "unsupported-schema"
  | "invalid-data"
  | "network";

export class ReportLoadError extends Error {
  constructor(
    public readonly kind: ReportLoadErrorKind,
    message: string,
    public readonly status?: number,
  ) {
    super(message);
    this.name = "ReportLoadError";
  }
}

const isObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

export function validateReportDocument(value: unknown): BmesReportDocument {
  if (!isObject(value)) {
    throw new ReportLoadError("invalid-data", "The report response is not a JSON object.");
  }

  if (value.contractId !== CONTRACT_ID) {
    throw new ReportLoadError("invalid-data", "Unsupported report contract.");
  }

  if (value.schemaVersion !== SUPPORTED_SCHEMA_VERSION) {
    throw new ReportLoadError(
      "unsupported-schema",
      `Unsupported report schema. Server ${String(value.schemaVersion)}, viewer ${SUPPORTED_SCHEMA_VERSION}`,
    );
  }

  if (!isObject(value.viewerDefaults) || !isObject(value.status) || !isObject(value.tabs)) {
    throw new ReportLoadError("invalid-data", "Required report fields are missing.");
  }

  const actualKeys = Object.keys(value.tabs).sort();
  const expectedKeys = [...TAB_KEYS].sort();
  if (
    actualKeys.length !== expectedKeys.length ||
    actualKeys.some((key, index) => key !== expectedKeys[index])
  ) {
    throw new ReportLoadError("invalid-data", "The eight-tab contract does not match.");
  }

  for (const key of TAB_KEYS) {
    const envelope = value.tabs[key];
    if (!isObject(envelope) || !isObject(envelope.status) || !("data" in envelope)) {
      throw new ReportLoadError("invalid-data", `Data for the ${key} tab is invalid.`);
    }
  }

  return value as unknown as BmesReportDocument;
}

export async function fetchReport(
  reportUrl: string,
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch,
): Promise<BmesReportDocument> {
  let response: Response;
  try {
    response = await fetcher(reportUrl, {
      method: "GET",
      credentials: "same-origin",
      cache: "no-store",
      headers: { Accept: "application/json" },
      signal,
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      throw error;
    }
    throw new ReportLoadError("network", "Cannot reach the report server.");
  }

  if (!response.ok) {
    if (response.status === 401) {
      throw new ReportLoadError("unauthorized", "Your sign-in session has expired.", 401);
    }
    if (response.status === 403) {
      throw new ReportLoadError("forbidden", "You do not have permission to view this report.", 403);
    }
    if (response.status === 404 || response.status === 410) {
      throw new ReportLoadError("expired", "The report link has expired or does not exist.", response.status);
    }
    throw new ReportLoadError("network", `Could not load the report. (HTTP ${response.status})`, response.status);
  }

  let payload: unknown;
  try {
    payload = await response.json();
  } catch {
    throw new ReportLoadError("invalid-data", "The report JSON could not be read.");
  }
  return validateReportDocument(payload);
}
