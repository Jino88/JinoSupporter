import { loginUrl } from "../../api";
import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { Entry, EntryFields, EntryLog, GroupsResponse, LogFields, ModelInfo, WorkingImage, WorkingModel } from "./types";

const base = "working";
const model = (name: string) => `${base}/models/${encodeURIComponent(name)}`;
const entry = (name: string, entryId: string) => `${model(name)}/entries/${encodeURIComponent(entryId)}`;
const logs = (name: string, entryId: string) => `${entry(name, entryId)}/logs`;

/** Multipart POST (toolSend is JSON only). */
async function upload<T>(path: string, form: FormData): Promise<T> {
  const response = await fetch(toolPath(path), {
    method: "POST",
    credentials: "same-origin",
    headers: { Accept: "application/json", "X-Requested-With": "PortalGNLAB" },
    body: form,
  });
  if (response.status === 401) window.location.assign(loginUrl(location.pathname));
  if (!response.ok) {
    const error = await response.json().catch(() => null);
    throw new Error(error?.error ?? `Request failed (HTTP ${response.status})`);
  }
  return (await response.json()) as T;
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
  /** Pictures for the issue itself (logId "") or for one of its notes. */
  addImages: async (name: string, entryId: string, logId: string, files: File[]) => {
    const form = new FormData();
    if (logId) form.append("logId", logId);
    files.forEach((f, i) => form.append("files", f, f.name || `pasted-${i + 1}.png`));
    return (await upload<{ images: WorkingImage[] }>(`${entry(name, entryId)}/images`, form)).images;
  },
  deleteImage: (name: string, entryId: string, imageId: string) =>
    toolSend<void>(`${entry(name, entryId)}/images/${encodeURIComponent(imageId)}`, undefined, "DELETE"),
  imageUrl: (imageId: string) => toolPath(`${base}/images/${encodeURIComponent(imageId)}`),
};
