import { ApiError, loginUrl } from "./api";
export type Values = Record<string, string | number | boolean | string[] | string[][]>;
export async function toolRequest<T>(path: string, body?: unknown, signal?: AbortSignal, onEvent?: (event: unknown) => void): Promise<T> {
 const form = body instanceof FormData;
 const response = await fetch("/api/portal/admin/" + path, { credentials: "same-origin", method: body === undefined ? "GET" : "POST", signal,
 headers: { Accept: "application/json", ...(body === undefined ? {} : { "X-Requested-With": "PortalGNLAB", ...(form ? {} : { "Content-Type": "application/json" }) }) },
 body: body === undefined ? undefined : form ? body : JSON.stringify(body) });
 if (response.status === 401) { window.location.assign(loginUrl(window.location.pathname)); throw new ApiError(401, "Signed out"); }
 if (!response.ok) { const error = await response.json().catch(() => null); throw new ApiError(response.status, error?.error ?? `Request failed (HTTP ${response.status})`); }
 if (response.headers.get("Content-Type")?.includes("application/x-ndjson")) {
  const reader = response.body!.getReader(); const decoder = new TextDecoder(); let pending = ""; const events: unknown[] = [];
  const emit = (line: string) => { if (line.trim()) { const event: unknown = JSON.parse(line); events.push(event); onEvent?.(event); } };
  try { while (true) { const {value,done} = await reader.read(); pending += decoder.decode(value, {stream: !done}); const lines = pending.split("\n"); pending = lines.pop()!; lines.forEach(emit); if (done) break; } emit(pending); }
  finally { reader.releaseLock(); }
  return events as T;
 }
 return response.json() as Promise<T>;
}
