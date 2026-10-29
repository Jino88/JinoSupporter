import { ApiError, getJson, loginUrl, send } from "./api";
import type { Memo, MemoView } from "./pages/teamMemo/types";
const root = "/api/portal/memo";
export const memoApi = {
    view: () => getJson<MemoView>(root), memo: (id: string) => getJson<Memo>(`${root}/items/${encodeURIComponent(id)}`),
    send: (method: "POST" | "PUT" | "DELETE", path: string, body?: unknown) => send(method, root + path, body),
    get: <T,>(path: string) => getJson<T>(root + path),
    async json<T>(path: string, body: unknown): Promise<T> { return request<T>(path, body); },
    async upload(path: string, form: FormData): Promise<void> { await request(path, form); },
    attachment: (id: string, download = false) => `${root}/attachments/${encodeURIComponent(id)}${download ? "?download=true" : ""}`,
};
async function request<T>(path: string, body: unknown): Promise<T> {
    const form = body instanceof FormData;
    const response = await fetch(root + path, { method: "POST", credentials: "same-origin", headers: { Accept: "application/json", "X-Requested-With": "PortalGNLAB", ...(form ? {} : { "Content-Type": "application/json" }) }, body: form ? body : JSON.stringify(body) });
    if (response.status === 401) {
        window.location.assign(loginUrl(window.location.pathname));
        throw new ApiError(401, "Signed out");
    }
    if (!response.ok) {
        const error = await response.json().catch(() => null);
        throw new ApiError(response.status, error?.error ?? `Request failed (HTTP ${response.status})`);
    }
    const text = await response.text();
    return (text ? JSON.parse(text) : null) as T;
}
