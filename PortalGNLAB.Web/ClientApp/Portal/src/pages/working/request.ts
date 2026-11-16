// Working > Request: a PROCESSING REQUEST (JIG/MC made, repaired or scrapped by the Manufacture Department). AI fills the
// form from a list pasted as text or pictures; the requests are kept on the server.
import { getJson, toolPath, toolSend } from "../../toolsApi";

export const REQUEST_KINDS = ["JIG/MC"] as const;
export const METHODS = [["new", "New"], ["repair", "Repair"], ["scrap", "Scrap (Báo phế)"]] as const;
export type Method = (typeof METHODS)[number][0];

export interface RequestItem {
  no: number;
  model: string;
  name: string;
  lineSpec: number;
  stock: number;
  order: number;
  drawingNo: string;
  picture: string;
  note: string;
}

/** What a request is saved with (the server adds id, who and when). */
export interface RequestFields {
  kind: string;
  requester: string;
  phone: string;
  checker: string;
  confirm: string;
  method: Method;
  departmentName: string;
  equipmentName: string;
  partName: string;
  itemApplication: string;
  usingPlace: string;
  requestDate: string;
  finishDate1: string;
  finishDate2: string;
  purpose: string;
  others: string;
  deptCode: string;
  qty1: string;
  qty2: string;
  items: RequestItem[];
}

export interface WorkingRequest extends RequestFields {
  id: string;
  createdBy: string;
  createdAt: string;
  updatedBy: string;
  updatedAt: string;
}

export const blankItem = (no: number, model = ""): RequestItem =>
  ({ no, model, name: "", lineSpec: 0, stock: 0, order: 0, drawingNo: "", picture: "", note: "" });

export const blankRequest = (today: string, kind = "JIG/MC"): RequestFields => ({
  kind, requester: "", phone: "", checker: "", confirm: "", method: "new", departmentName: "ME GROUP",
  equipmentName: "", partName: "", itemApplication: "", usingPlace: "", requestDate: today, finishDate1: "", finishDate2: "",
  purpose: "New request", others: "", deptCode: "", qty1: "", qty2: "", items: [],
});

const text = (v: unknown) => (typeof v === "string" ? v : typeof v === "number" && Number.isFinite(v) ? String(v) : "");
/** A count: numbers and number text ("40", "1,200") read, anything else or below 0 is 0. */
export const count = (v: unknown): number => {
  const n = typeof v === "number" ? v : Number(text(v).replace(/,/g, "").trim() || NaN);
  return Number.isFinite(n) && n > 0 ? n : 0;
};

/** yyyy-MM-dd from yyyy-MM-dd, yyyy.MM.dd or yyyy/M/d; anything else "". */
export function isoDate(v: unknown): string {
  const m = /^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/.exec(text(v).trim());
  if (!m) return "";
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const t = new Date(Date.UTC(y, mo - 1, d));
  if (t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return "";
  return `${m[1]}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

/** The date as the form writes it: 2026.11.05. */
export const dotDate = (iso: string) => (isoDate(iso) ? isoDate(iso).replace(/-/g, ".") : iso);

/** The order of a row: as given, else what the line needs beyond the stock. */
export const orderOf = (i: Pick<RequestItem, "lineSpec" | "stock" | "order">) => (i.order > 0 ? i.order : Math.max(0, i.lineSpec - i.stock));

/** Rows with a name, numbered from 1. */
export const keptItems = (items: RequestItem[]): RequestItem[] =>
  items.filter((i) => i.name.trim()).map((i, n) => ({ ...i, no: n + 1 }));

/** A request (saved or answered by AI) read leniently. */
export function readRequest(v: unknown, today = ""): RequestFields & Partial<WorkingRequest> {
  const r = (v && typeof v === "object" ? v : {}) as Record<string, unknown>;
  const base = blankRequest(today);
  // a field left out takes the form's default (ME GROUP, New request); one sent empty stays empty
  const s = (k: keyof RequestFields) => (r[k] === undefined || r[k] === null ? (base[k] as string) : text(r[k]).trim());
  const method = METHODS.some(([m]) => m === r.method) ? (r.method as Method) : "new";
  const items = (Array.isArray(r.items) ? r.items : []).map((x, n): RequestItem => {
    const i = (x && typeof x === "object" ? x : {}) as Record<string, unknown>;
    return {
      no: count(i.no) || n + 1, model: text(i.model).trim(), name: text(i.name).trim(),
      lineSpec: count(i.lineSpec), stock: count(i.stock), order: count(i.order),
      drawingNo: text(i.drawingNo).trim(), picture: text(i.picture).trim(), note: text(i.note).trim(),
    };
  });
  const out: RequestFields & Partial<WorkingRequest> = {
    kind: s("kind"), requester: s("requester"), phone: s("phone"), checker: s("checker"), confirm: s("confirm"), method,
    departmentName: s("departmentName"), equipmentName: s("equipmentName"), partName: s("partName"),
    itemApplication: s("itemApplication"), usingPlace: s("usingPlace"),
    requestDate: r.requestDate === undefined ? base.requestDate : isoDate(r.requestDate),
    finishDate1: isoDate(r.finishDate1), finishDate2: isoDate(r.finishDate2),
    purpose: s("purpose"), others: s("others"), deptCode: s("deptCode"), qty1: s("qty1"), qty2: s("qty2"), items,
  };
  for (const k of ["id", "createdBy", "createdAt", "updatedBy", "updatedAt"] as const) if (typeof r[k] === "string") out[k] = r[k] as string;
  return out;
}

/** The fields to save: rows without a name dropped, the rest renumbered, an empty order made from line and stock. */
export const toSave = (r: RequestFields): RequestFields =>
  ({ ...requestFields(r), items: keptItems(r.items).map((i) => ({ ...i, order: orderOf(i) })) });

/** Only the form fields (no id or who saved it). */
export function requestFields(r: RequestFields): RequestFields {
  const out = {} as Record<string, unknown>;
  for (const k of Object.keys(blankRequest("")) as (keyof RequestFields)[]) out[k] = r[k];
  return out as unknown as RequestFields;
}

/** What is wrong before saving, or "". */
export function requestProblem(r: RequestFields): string {
  if (!r.partName.trim() && !r.equipmentName.trim()) return "Type the part name or the equipment name.";
  if (!keptItems(r.items).length) return "Add at least one item with a name.";
  return "";
}

/** The model most rows are for: the part name when AI or the user gave none. */
export function commonModel(items: RequestItem[]): string {
  const n = new Map<string, number>();
  for (const i of items) if (i.model.trim()) n.set(i.model.trim(), (n.get(i.model.trim()) ?? 0) + 1);
  return [...n.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "";
}

/** Totals of the rows: line spec, stock and order. */
export const totals = (items: RequestItem[]) => items.reduce(
  (t, i) => ({ lineSpec: t.lineSpec + i.lineSpec, stock: t.stock + i.stock, order: t.order + orderOf(i) }),
  { lineSpec: 0, stock: 0, order: 0 });

/** A saved request as one line of the list: dept code, part, date and how many rows. */
export const requestTitle = (r: RequestFields) =>
  [r.deptCode || "(no code)", r.partName || r.equipmentName || "(no part)"].join(" · ");

const base = "working/requests";
const one = (id: string) => `${base}/${encodeURIComponent(id)}`;

export const requestApi = {
  list: () => getJson<{ requests: unknown[] }>(toolPath(base)).then((x) => (x.requests ?? []).map((r) => readRequest(r) as WorkingRequest)),
  create: (f: RequestFields) => toolSend<unknown>(base, toSave(f)).then((r) => readRequest(r) as WorkingRequest),
  save: (id: string, f: RequestFields) => toolSend<unknown>(one(id), toSave(f), "PUT").then((r) => readRequest(r) as WorkingRequest),
  remove: (id: string) => toolSend<void>(one(id), undefined, "DELETE"),
  /** The form filled by AI from what was pasted (nothing saved). */
  ai: (body: { text: string; today: string; kind: string; current: RequestFields; images: { name: string; dataUrl: string }[] }, signal?: AbortSignal) =>
    toolSend<{ request: unknown; note: string; model: string }>(`${base}/ai`, { ...body, current: toSave(body.current) }, "POST", signal),
};
