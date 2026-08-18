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
    throw new ReportLoadError("invalid-data", "리포트 응답이 JSON 객체가 아닙니다.");
  }

  if (value.contractId !== CONTRACT_ID) {
    throw new ReportLoadError("invalid-data", "지원하지 않는 리포트 계약입니다.");
  }

  if (value.schemaVersion !== SUPPORTED_SCHEMA_VERSION) {
    throw new ReportLoadError(
      "unsupported-schema",
      `지원하지 않는 리포트 스키마입니다. 서버 ${String(value.schemaVersion)}, 뷰어 ${SUPPORTED_SCHEMA_VERSION}`,
    );
  }

  if (!isObject(value.viewerDefaults) || !isObject(value.status) || !isObject(value.tabs)) {
    throw new ReportLoadError("invalid-data", "리포트 공통 필드가 누락되었습니다.");
  }

  const actualKeys = Object.keys(value.tabs).sort();
  const expectedKeys = [...TAB_KEYS].sort();
  if (
    actualKeys.length !== expectedKeys.length ||
    actualKeys.some((key, index) => key !== expectedKeys[index])
  ) {
    throw new ReportLoadError("invalid-data", "8개 탭 계약이 일치하지 않습니다.");
  }

  for (const key of TAB_KEYS) {
    const envelope = value.tabs[key];
    if (!isObject(envelope) || !isObject(envelope.status) || !("data" in envelope)) {
      throw new ReportLoadError("invalid-data", `${key} 탭 데이터가 올바르지 않습니다.`);
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
    throw new ReportLoadError("network", "리포트 서버에 연결할 수 없습니다.");
  }

  if (!response.ok) {
    if (response.status === 401) {
      throw new ReportLoadError("unauthorized", "로그인 세션이 만료되었습니다.", 401);
    }
    if (response.status === 403) {
      throw new ReportLoadError("forbidden", "이 리포트를 볼 권한이 없습니다.", 403);
    }
    if (response.status === 404 || response.status === 410) {
      throw new ReportLoadError("expired", "리포트 링크가 만료되었거나 존재하지 않습니다.", response.status);
    }
    throw new ReportLoadError("network", `리포트를 불러오지 못했습니다. (HTTP ${response.status})`, response.status);
  }

  let payload: unknown;
  try {
    payload = await response.json();
  } catch {
    throw new ReportLoadError("invalid-data", "리포트 JSON을 읽을 수 없습니다.");
  }
  return validateReportDocument(payload);
}
