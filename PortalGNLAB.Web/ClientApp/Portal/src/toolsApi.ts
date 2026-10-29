import { getJson, loginUrl } from "./api";
export { getJson };

export const toolPath = (tool: string) => "/api/portal/tools/" + tool;

export async function toolSend<T>(
  path: string,
  body?: unknown,
  method = "POST",
  signal?: AbortSignal,
): Promise<T> {
  const response = await fetch(toolPath(path), {
    method,
    signal,
    credentials: "same-origin",
    headers: {
      "Content-Type": "application/json",
      "X-Requested-With": "PortalGNLAB",
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  if (response.status === 401)
    window.location.assign(loginUrl(location.pathname));
  if (!response.ok) {
    const error = await response.json().catch(() => null);
    throw new Error(error?.error ?? `Request failed (HTTP ${response.status})`);
  }
  if (response.status === 204) return undefined as T;
  if (response.headers.get("content-type")?.includes("spreadsheet"))
    return (await response.blob()) as T;
  return (await response.json()) as T;
}

export function imageData(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1]);
    reader.onerror = () => reject(new Error("Image could not be read."));
    reader.readAsDataURL(file);
  });
}

let legacyScripts: Promise<void> | undefined;
/** Same locally shipped Chart.js and renderer as Blazor; no additional npm dependency. */
export function loadLegacyTools(): Promise<void> {
  if (!legacyScripts) {
    legacyScripts = (async () => {
      for (const src of ["/js/chart.umd.min.js", "/js/app.js"]) {
        await new Promise<void>((resolve, reject) => {
          const script = document.createElement("script");
          script.src = src;
          script.onload = () => resolve();
          script.onerror = () => reject(new Error(`Could not load ${src}`));
          document.head.appendChild(script);
        });
      }
    })().catch((error) => {
      legacyScripts = undefined;
      throw error;
    });
  }
  return legacyScripts;
}

export function downloadBlob(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}
