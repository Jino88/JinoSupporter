import { getJson, loginUrl } from "./api";
export { getJson };
export const toolPath = (tool: string) => "/api/portal/tools/" + tool;
export async function toolSend<T>(path: string, body?: unknown, method = "POST"): Promise<T> {
  const response = await fetch(toolPath(path), { method, credentials: "same-origin", headers: { "Content-Type": "application/json", "X-Requested-With": "PortalGNLAB" }, body: body === undefined ? undefined : JSON.stringify(body) });
  if (response.status === 401) window.location.assign(loginUrl(location.pathname));
  if (!response.ok) { const error = await response.json().catch(() => null); throw new Error(error?.error ?? `Request failed (HTTP ${response.status})`); }
  if (response.status === 204) return undefined as T;
  if (response.headers.get("content-type")?.includes("spreadsheet")) return await response.blob() as T;
  return await response.json() as T;
}
export async function imageData(file: File): Promise<string> {
  return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = () => resolve(String(reader.result).split(",")[1]); reader.onerror = () => reject(new Error("Image could not be read.")); reader.readAsDataURL(file); });
}
