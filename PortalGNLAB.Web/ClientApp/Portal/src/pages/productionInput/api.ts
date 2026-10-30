import { getJson, toolPath, toolSend } from "../../toolsApi";
import type { Day, DayData, DayStatus, DaySummary, Model, ModelSetup, ModelSummary, Plan, PlanData, Sheet } from "./types";

const base = "production-input";
const model = (id: string) => `${base}/models/${encodeURIComponent(id)}`;

export const productionApi = {
  models: () => getJson<{ models: ModelSummary[] }>(toolPath(base)).then((r) => r.models),
  model: (id: string) => getJson<Model>(toolPath(model(id))),
  createModel: (body: { name: string; build: string; setup: ModelSetup }) => toolSend<Model>(`${base}/models`, body),
  saveModel: (m: Model) =>
    toolSend<Model>(model(m.id), { name: m.name, build: m.build, setup: m.setup, revision: m.revision }, "PUT"),
  deleteModel: (id: string) => toolSend<void>(model(id), undefined, "DELETE"),
  day: (id: string, date: string) => getJson<{ day: Day | null; previous: Day | null }>(toolPath(`${model(id)}/days/${date}`)),
  days: (id: string, from: string, to: string) =>
    getJson<Day[]>(toolPath(`${model(id)}/days?from=${from}&to=${to}`)),
  saveDay: (id: string, date: string, body: { status: DayStatus; data: DayData; summary: DaySummary; revision: number }) =>
    toolSend<Day>(`${model(id)}/days/${date}`, body, "PUT"),
  plan: (id: string) => getJson<Plan>(toolPath(`${model(id)}/plan`)),
  savePlan: (id: string, data: PlanData, revision: number) => toolSend<Plan>(`${model(id)}/plan`, { data, revision }, "PUT"),
  exportXlsx: (fileName: string, sheets: Sheet[]) => toolSend<Blob>(`${base}/export`, { fileName, sheets }),
};
