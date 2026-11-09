import { IMAGE_TYPES, MAX_IMAGE_BYTES, MAX_IMAGES_PER_UPLOAD } from "./types";
import type { Entry, EntryFields, EntryKind, EntryLog, LogFields, ModelGroup } from "./types";

export interface GroupNode {
  id: string;
  name: string;
  models: string[];
  children: GroupNode[];
  /** Models in this group and every group under it. */
  total: number;
}

const byOrder = (a: ModelGroup, b: ModelGroup) => a.sort - b.sort || a.position - b.position || a.name.localeCompare(b.name);

/** model_groups rows → tree. A row whose parent is missing becomes a root; cycles are cut. */
export function buildTree(rows: ModelGroup[]): GroupNode[] {
  const ids = new Set(rows.map((r) => r.id));
  const kids = new Map<string | null, ModelGroup[]>();
  for (const r of rows) {
    const parent = r.parentId && ids.has(r.parentId) && r.parentId !== r.id ? r.parentId : null;
    kids.set(parent, [...(kids.get(parent) ?? []), r]);
  }
  const seen = new Set<string>();
  const make = (r: ModelGroup): GroupNode | null => {
    if (seen.has(r.id)) return null;
    seen.add(r.id);
    const children = (kids.get(r.id) ?? []).sort(byOrder).map(make).filter((n): n is GroupNode => n !== null);
    const models = [...r.models];
    return { id: r.id, name: r.name, models, children, total: models.length + children.reduce((s, c) => s + c.total, 0) };
  };
  const roots = (kids.get(null) ?? []).sort(byOrder).map(make).filter((n): n is GroupNode => n !== null);
  // rows only reachable through a cycle were never visited; show them as roots
  for (const r of rows.filter((x) => !seen.has(x.id)).sort(byOrder)) {
    const n = make(r);
    if (n) roots.push(n);
  }
  return roots;
}

/** Tree limited to models whose name contains the query (case-insensitive); empty groups drop out. */
export function filterTree(nodes: GroupNode[], query: string): GroupNode[] {
  const q = query.trim().toLowerCase();
  if (!q) return nodes;
  return nodes.flatMap((n) => {
    const models = n.models.filter((m) => m.toLowerCase().includes(q));
    const children = filterTree(n.children, q);
    return models.length || children.length ? [{ ...n, models, children }] : [];
  });
}

/** Path of group names down to the model, e.g. ["MODULE", "MSM"]; [] when the model is not in the tree. */
export function pathOf(nodes: GroupNode[], model: string, trail: string[] = []): string[] {
  for (const n of nodes) {
    if (n.models.includes(model)) return [...trail, n.name];
    const found = pathOf(n.children, model, [...trail, n.name]);
    if (found.length) return found;
  }
  return [];
}

export function firstModel(nodes: GroupNode[]): string | null {
  for (const n of nodes) {
    if (n.models.length) return n.models[0];
    const m = firstModel(n.children);
    if (m) return m;
  }
  return null;
}

/** Newest first; same date → newest created first. */
export function sortEntries(list: Entry[]): Entry[] {
  return [...list].sort((a, b) => b.date.localeCompare(a.date) || b.createdAt.localeCompare(a.createdAt));
}

export function countByKind(list: Entry[]): Record<EntryKind, number> {
  const c: Record<EntryKind, number> = { dev: 0, issue: 0, review: 0 };
  for (const e of list) c[e.kind]++;
  return c;
}

/** An issue is closed once it has a done date (or was closed before done dates existed). */
export const isOpenIssue = (e: Entry) => e.kind === "issue" && !e.doneDate && e.status !== "Closed";

/** An open issue whose due date has passed. */
export const isLate = (e: Entry, today: string) => isOpenIssue(e) && !!e.due && e.due < today;

/** Entries grouped under their yyyy-MM month, newest month first (input must be sorted). */
export function byMonth(list: Entry[]): { month: string; entries: Entry[] }[] {
  const out: { month: string; entries: Entry[] }[] = [];
  for (const e of list) {
    const month = e.date.slice(0, 7);
    const last = out.at(-1);
    if (last && last.month === month) last.entries.push(e);
    else out.push({ month, entries: [e] });
  }
  return out;
}

export function blankEntry(kind: EntryKind, today: string, phase = ""): EntryFields {
  return {
    kind, date: today, title: "", body: "", owner: "", due: "",
    status: kind === "issue" ? "Open" : "",
    severity: "",
    result: kind === "review" ? "Pass" : "",
    phase: kind === "issue" ? phase : "",
    category: "", process: "", cause: "", doneDate: "", finalResult: "",
  };
}

export const fieldsOf = (e: Entry): EntryFields => ({
  kind: e.kind, date: e.date, title: e.title, body: e.body, status: e.status,
  severity: e.severity, result: e.result, owner: e.owner, due: e.due,
  phase: e.phase ?? "", category: e.category ?? "", process: e.process ?? "", cause: e.cause ?? "",
  doneDate: e.doneDate ?? "", finalResult: e.finalResult ?? "",
});

/** Issue fields with Status kept in step with the done date. */
export const issueFields = (f: EntryFields): EntryFields => ({ ...f, status: f.doneDate ? "Closed" : "Open" });

/** Issues in number order (No. 1 first). */
export const issuesBySeq = (list: Entry[]) => list.filter((e) => e.kind === "issue").sort((a, b) => a.seq - b.seq || a.createdAt.localeCompare(b.createdAt));

/** Values already used in a field, followed by the suggestions, without duplicates (case-insensitive). */
export function suggestions(used: string[], defaults: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const v of [...defaults, ...used]) {
    const t = v.trim();
    if (t && !seen.has(t.toLowerCase())) { seen.add(t.toLowerCase()); out.push(t); }
  }
  return out;
}

export const blankLog = (today: string): LogFields => ({ date: today, text: "", owner: "", due: "", doneDate: "" });

export const logFieldsOf = (l: EntryLog): LogFields =>
  ({ date: l.date, text: l.text, owner: l.owner ?? "", due: l.due ?? "", doneDate: l.doneDate ?? "" });

/** Notes in date order, oldest first. */
export const notesInOrder = (logs: EntryLog[]) =>
  [...logs].sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));

/** A note that is not done and whose due date has passed. */
export const isLogLate = (l: LogFields, today: string) => !l.doneDate && !!l.due && l.due < today;

export function logProblem(f: LogFields): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) return "Date must be yyyy-MM-dd.";
  if (!f.text.trim()) return "Write what was done.";
  if (f.text.length > 5000) return "A note must be 5000 characters or fewer.";
  if (f.owner.trim().length > 100) return "Owner must be 100 characters or fewer.";
  if (f.due && !/^\d{4}-\d{2}-\d{2}$/.test(f.due)) return "Due must be yyyy-MM-dd.";
  if (f.doneDate && !/^\d{4}-\d{2}-\d{2}$/.test(f.doneDate)) return "Done date must be yyyy-MM-dd.";
  if (f.doneDate && f.doneDate < f.date) return "Done date cannot be before the note date.";
  return "";
}

/**
 * Pictures among pasted / picked files. Other files are ignored so a normal text paste still works;
 * a picture the server would refuse comes back as a problem instead.
 */
export function pastedImages<F extends { name: string; type: string; size: number }>(files: F[], have: number): { files: F[]; problem: string } {
  const pictures = files.filter((f) => f.type.startsWith("image/"));
  const refused = pictures.find((f) => !IMAGE_TYPES.includes(f.type));
  if (refused) return { files: [], problem: `${refused.name || "This picture"} is not PNG, JPEG, GIF, WebP or BMP.` };
  const large = pictures.find((f) => f.size > MAX_IMAGE_BYTES);
  if (large) return { files: [], problem: `${large.name || "This picture"} is larger than 10 MB.` };
  if (have + pictures.length > MAX_IMAGES_PER_UPLOAD) return { files: [], problem: `Paste at most ${MAX_IMAGES_PER_UPLOAD} pictures at a time.` };
  return { files: pictures, problem: "" };
}

/** Message for a form that cannot be saved yet, or "" when it can. */
export function entryProblem(f: EntryFields): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) return "Date must be yyyy-MM-dd.";
  if (!f.title.trim()) return "Title is required.";
  if (f.title.trim().length > 200) return "Title must be 200 characters or fewer.";
  if (f.due && !/^\d{4}-\d{2}-\d{2}$/.test(f.due)) return "Due must be yyyy-MM-dd.";
  if (f.doneDate && !/^\d{4}-\d{2}-\d{2}$/.test(f.doneDate)) return "Done date must be yyyy-MM-dd.";
  if (f.doneDate && f.doneDate < f.date) return "Done date cannot be before the registered date.";
  return "";
}
