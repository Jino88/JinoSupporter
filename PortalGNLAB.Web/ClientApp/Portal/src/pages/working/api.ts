import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { Entry, EntryFields, GroupsResponse, ModelInfo, WorkingModel } from "./types";

const base = "working";
const model = (name: string) => `${base}/models/${encodeURIComponent(name)}`;

export const workingApi = {
  groups: () => getJson<GroupsResponse>(toolPath(`${base}/groups`)),
  model: (name: string) => getJson<WorkingModel>(toolPath(model(name))),
  saveModel: (name: string, info: ModelInfo, revision: number) => toolSend<WorkingModel>(model(name), { info, revision }, "PUT"),
  entries: (name: string) => getJson<{ entries: Entry[] }>(toolPath(`${model(name)}/entries`)).then((r) => r.entries),
  addEntry: (name: string, fields: EntryFields) => toolSend<Entry>(`${model(name)}/entries`, fields),
  saveEntry: (name: string, id: string, fields: EntryFields, revision: number) =>
    toolSend<Entry>(`${model(name)}/entries/${encodeURIComponent(id)}`, { ...fields, revision }, "PUT"),
  deleteEntry: (name: string, id: string) => toolSend<void>(`${model(name)}/entries/${encodeURIComponent(id)}`, undefined, "DELETE"),
};
