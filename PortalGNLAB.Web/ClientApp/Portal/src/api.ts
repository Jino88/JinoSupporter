// JSON API of PortalGNLAB.Web for the React app (Endpoints/PortalShellEndpoints.cs).
// Sign-in uses the same cookie as the Blazor pages; a 401 means "go to /login".

export interface PortalUser {
  name: string;
  role: string;
}

export interface MenuItem {
  id: string;
  label: string;
  group: string;
  /** Blazor route, or /app/... once the page is ported to React. */
  href: string;
  react: boolean;
}

export interface DownloadAsset {
  fileName: string;
  sizeBytes: number;
  sizeText: string;
  url: string;
}

export interface PcDownload {
  version: string;
  notes: string;
  publishedAt: string | null;
  setup: DownloadAsset | null;
  package: DownloadAsset | null;
}

export class ApiError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
  }
}

/** Login page URL that brings the user back to where they were. */
export function loginUrl(returnTo: string): string {
  return "/login?returnUrl=" + encodeURIComponent(returnTo);
}

async function getJson<T>(path: string): Promise<T> {
  const response = await fetch(path, { credentials: "same-origin", headers: { Accept: "application/json" } });
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    throw new ApiError(401, "Signed out");
  }
  if (!response.ok) throw new ApiError(response.status, `Request failed (HTTP ${response.status})`);
  const text = await response.text();
  return (text ? JSON.parse(text) : null) as T;
}

export const api = {
  me: () => getJson<PortalUser>("/api/portal/me"),
  menu: () => getJson<MenuItem[]>("/api/portal/menu"),
  pcDownload: () => getJson<PcDownload | null>("/api/portal/pc-download"),
};
