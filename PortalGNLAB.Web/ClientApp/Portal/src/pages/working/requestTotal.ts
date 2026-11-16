// Working > Request Total: how many of each JIG were asked for in the saved requests, model by model. A JIG is its name and
// its drawing number together: the same name with another drawing number (or none) is another JIG.
import { drawingKey, orderOf, type Method, type WorkingRequest } from "./request";

export interface JigTotal {
  /** name and drawing number as compared (the key of the JIG) */
  key: string;
  /** as first written */
  name: string;
  drawingNo: string;
  /** the order asked for, by model */
  byModel: Record<string, number>;
  total: number;
  /** how many requests asked for it */
  requests: number;
}

export interface ModelTotal { model: string; total: number; jigs: number; requests: number }

export interface RequestTotals {
  /** every model, most ordered first */
  models: ModelTotal[];
  /** every JIG, those of the most models first, then the most ordered */
  jigs: JigTotal[];
  /** the requests counted */
  requests: number;
}

export interface TotalFilter {
  /** yyyy-MM-dd, request date from / to (inclusive); "" for no limit */
  from?: string;
  to?: string;
  /** "" for every processing method */
  method?: Method | "";
}

export const NO_MODEL = "(no model)";

/** A JIG name compared without case or extra spaces. */
export const jigName = (name: string) => name.trim().replace(/\s+/g, " ").toUpperCase();
/** The key of a JIG: its name and its drawing number both. */
export const jigKey = (name: string, drawingNo: string) => `${jigName(name)}\u0000${drawingKey(drawingNo)}`;
/** The model of a row: its own, else the part name of the request. */
const modelOf = (rowModel: string, r: WorkingRequest) => rowModel.trim() || r.partName.trim() || NO_MODEL;

/** The requests kept by the filter: request date in the range, the method chosen. */
export const filtered = (list: WorkingRequest[], f: TotalFilter = {}) =>
  list.filter((r) => (!f.from || (r.requestDate && r.requestDate >= f.from)) && (!f.to || (r.requestDate && r.requestDate <= f.to))
    && (!f.method || r.method === f.method));

/** The order of every JIG by model and of every model, from the requests. Rows without a name are left out. */
export function requestTotals(list: WorkingRequest[]): RequestTotals {
  const jigs = new Map<string, JigTotal & { asked: Set<string> }>();
  const models = new Map<string, { total: number; jigs: Set<string>; asked: Set<string> }>();
  list.forEach((r, n) => {
    const req = r.id || `#${n}`;
    for (const i of r.items) {
      if (!jigName(i.name)) continue;
      const key = jigKey(i.name, i.drawingNo);
      const model = modelOf(i.model, r);
      const qty = orderOf(i);
      const j = jigs.get(key) ?? { key, name: i.name.trim(), drawingNo: i.drawingNo.trim(), byModel: {}, total: 0, requests: 0, asked: new Set<string>() };
      j.byModel[model] = (j.byModel[model] ?? 0) + qty;
      j.total += qty;
      j.asked.add(req);
      jigs.set(key, j);
      const m = models.get(model) ?? { total: 0, jigs: new Set<string>(), asked: new Set<string>() };
      m.total += qty;
      m.jigs.add(key);
      m.asked.add(req);
      models.set(model, m);
    }
  });
  return {
    models: [...models.entries()]
      .map(([model, m]) => ({ model, total: m.total, jigs: m.jigs.size, requests: m.asked.size }))
      .sort((a, b) => b.total - a.total || a.model.localeCompare(b.model)),
    jigs: [...jigs.values()]
      .map(({ asked, ...j }) => ({ ...j, requests: asked.size }))
      .sort((a, b) => modelCount(b) - modelCount(a) || b.total - a.total || a.name.localeCompare(b.name) || a.drawingNo.localeCompare(b.drawingNo)),
    requests: list.length,
  };
}

/** How many models asked for the JIG. */
export const modelCount = (j: JigTotal) => Object.keys(j.byModel).length;

/** A JIG common to models: asked for by two or more. */
export const isCommon = (j: JigTotal) => modelCount(j) >= 2;

/** The JIGs whose name or drawing number has the text (any case). */
export const searched = (jigs: JigTotal[], q: string) => {
  const t = q.trim().toUpperCase();
  return t ? jigs.filter((j) => j.name.toUpperCase().includes(t) || j.drawingNo.toUpperCase().includes(t)) : jigs;
};
