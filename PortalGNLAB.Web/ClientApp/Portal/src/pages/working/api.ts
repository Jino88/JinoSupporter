import { loginUrl } from "../../api";
import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { placeRequest } from "./aiPlace";
import type { ManualGroup } from "./manual";
import type { Entry, EntryFields, EntryLog, GroupsResponse, LogFields, ModelInfo, WorkingImage, WorkingModel } from "./types";

const base = "working";
const model = (name: string) => `${base}/models/${encodeURIComponent(name)}`;
const entry = (name: string, entryId: string) => `${model(name)}/entries/${encodeURIComponent(entryId)}`;
const logs = (name: string, entryId: string) => `${entry(name, entryId)}/logs`;

/** Multipart POST (toolSend is JSON only); XMLHttpRequest so a long video upload can show how far it is (0–1). */
function upload<T>(path: string, form: FormData, onProgress?: (done: number) => void): Promise<T> {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", toolPath(path));
    xhr.withCredentials = true;
    xhr.setRequestHeader("Accept", "application/json");
    xhr.setRequestHeader("X-Requested-With", "PortalGNLAB");
    xhr.upload.onprogress = (e) => { if (e.lengthComputable) onProgress?.(e.loaded / e.total); };
    xhr.onload = () => {
      if (xhr.status === 401) window.location.assign(loginUrl(location.pathname));
      let body: { error?: string } | null = null;
      try { body = JSON.parse(xhr.responseText); } catch { /* not JSON */ }
      if (xhr.status >= 200 && xhr.status < 300) resolve(body as T);
      else reject(new Error(body?.error ?? (xhr.status === 413 ? "The files are too large." : `Request failed (HTTP ${xhr.status})`)));
    };
    xhr.onerror = () => reject(new Error("The upload failed. Check the connection and try again."));
    xhr.send(form);
  });
}

export const workingApi = {
  groups: () => getJson<GroupsResponse>(toolPath(`${base}/groups`)),
  model: (name: string) => getJson<WorkingModel>(toolPath(model(name))),
  saveModel: (name: string, info: ModelInfo, revision: number) => toolSend<WorkingModel>(model(name), { info, revision }, "PUT"),
  entries: (name: string) => getJson<{ entries: Entry[] }>(toolPath(`${model(name)}/entries`)).then((r) => r.entries),
  addEntry: (name: string, fields: EntryFields) => toolSend<Entry>(`${model(name)}/entries`, fields),
  saveEntry: (name: string, id: string, fields: EntryFields, revision: number) =>
    toolSend<Entry>(entry(name, id), { ...fields, revision }, "PUT"),
  deleteEntry: (name: string, id: string) => toolSend<void>(entry(name, id), undefined, "DELETE"),
  addLog: (name: string, entryId: string, fields: LogFields) => toolSend<EntryLog>(logs(name, entryId), fields),
  saveLog: (name: string, entryId: string, log: EntryLog, fields: LogFields) =>
    toolSend<EntryLog>(`${logs(name, entryId)}/${encodeURIComponent(log.id)}`, { ...fields, revision: log.revision }, "PUT"),
  deleteLog: (name: string, entryId: string, logId: string) =>
    toolSend<void>(`${logs(name, entryId)}/${encodeURIComponent(logId)}`, undefined, "DELETE"),
  /** Pictures and videos for the issue itself (logId "") or for one of its notes, sent as they are. */
  addImages: async (name: string, entryId: string, logId: string, files: File[], onProgress?: (done: number) => void) => {
    const form = new FormData();
    if (logId) form.append("logId", logId);
    files.forEach((f, i) => form.append("files", f, f.name || `pasted-${i + 1}.png`));
    return (await upload<{ images: WorkingImage[] }>(`${entry(name, entryId)}/images`, form, onProgress)).images;
  },
  deleteImage: (name: string, entryId: string, imageId: string) =>
    toolSend<void>(`${entry(name, entryId)}/images/${encodeURIComponent(imageId)}`, undefined, "DELETE"),
  /** The hand-entered process table changed by AI as the request says (not saved; Save keeps it). */
  manualAi: (name: string, request: string, group: ManualGroup, signal?: AbortSignal) =>
    toolSend<{ group: unknown; note: string; model: string }>(`${model(name)}/manual/ai`, { request, group }, "POST", signal),
  /** Text placed by AI into Model Info, Dev Plan and new entries (aiPlace.ts); nothing is saved here. */
  placeAi: (name: string, request: ReturnType<typeof placeRequest>, signal?: AbortSignal) =>
    toolSend<unknown>(`${model(name)}/ai/place`, request, "POST", signal),
  /** Translations of texts ready on the server (en / ko / vi); the rest are translated there in the background. */
  translations: (lang: string, texts: string[]) =>
    toolSend<{ lang: string; translations: Record<string, string>; missing: number; running: boolean }>(`${base}/translations`, { lang, texts }),
  imageUrl: (imageId: string) => toolPath(`${base}/images/${encodeURIComponent(imageId)}`),
};
