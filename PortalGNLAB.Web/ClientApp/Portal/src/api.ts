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

export async function getJson<T>(path: string): Promise<T> {
  const response = await fetch(path, { credentials: "same-origin", headers: { Accept: "application/json" } });
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    throw new ApiError(401, "Signed out");
  }
  if (!response.ok) throw new ApiError(response.status, `Request failed (HTTP ${response.status})`);
  const text = await response.text();
  return (text ? JSON.parse(text) : null) as T;
}

/** POST / PUT / DELETE with a JSON body. The server answers {error} with 400 for rejected input. */
export async function send(method: "POST" | "PUT" | "DELETE", path: string, body?: unknown): Promise<void> {
  const response = await fetch(path, {
    method,
    credentials: "same-origin",
    headers: {
      Accept: "application/json",
      "Content-Type": "application/json",
      // Required by the server for changes (AdminUserEndpoints.RequestedWithValue).
      "X-Requested-With": "PortalGNLAB",
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (response.status === 401) {
    window.location.assign(loginUrl(window.location.pathname + window.location.search));
    throw new ApiError(401, "Signed out");
  }
  if (!response.ok) throw new ApiError(response.status, await errorText(response));
}

async function errorText(response: Response): Promise<string> {
  try {
    const data = (await response.json()) as { error?: string };
    if (data?.error) return data.error;
  } catch {
    // not JSON
  }
  return `Request failed (HTTP ${response.status})`;
}

export interface AdminUser {
  id: number;
  username: string;
  displayName: string;
  role: string;
  createdAt: string;
}

export interface AdminMenu {
  id: string;
  label: string;
  group: string;
}

export interface AdminUsersView {
  currentUser: string;
  roles: string[];
  users: AdminUser[];
  menus: AdminMenu[];
  /** Role → menu ids. The Admin role always has every menu. */
  permissions: Record<string, string[]>;
}

export const api = {
  me: () => getJson<PortalUser>("/api/portal/me"),
  menu: () => getJson<MenuItem[]>("/api/portal/menu"),
  pcDownload: () => getJson<PcDownload | null>("/api/portal/pc-download"),
  admin: {
    users: () => getJson<AdminUsersView>("/api/portal/admin/users"),
    addUser: (user: { username: string; displayName: string; password: string; role: string }) =>
      send("POST", "/api/portal/admin/users", user),
    setRole: (id: number, role: string) => send("PUT", `/api/portal/admin/users/${id}/role`, { role }),
    setName: (id: number, displayName: string) => send("PUT", `/api/portal/admin/users/${id}/name`, { displayName }),
    setPassword: (id: number, password: string) => send("PUT", `/api/portal/admin/users/${id}/password`, { password }),
    removeUser: (id: number) => send("DELETE", `/api/portal/admin/users/${id}`),
    setPermission: (role: string, menuId: string, allowed: boolean) =>
      send("PUT", "/api/portal/admin/permissions", { role, menuId, allowed }),
  },
};
