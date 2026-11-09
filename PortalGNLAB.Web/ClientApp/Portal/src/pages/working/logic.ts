import type { Entry, EntryFields, EntryKind, ModelGroup } from "./types";

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

export const isOpenIssue = (e: Entry) => e.kind === "issue" && e.status !== "Closed";

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

export function blankEntry(kind: EntryKind, today: string): EntryFields {
  return {
    kind, date: today, title: "", body: "", owner: "", due: "",
    status: kind === "issue" ? "Open" : "",
    severity: kind === "issue" ? "Major" : "",
    result: kind === "review" ? "Pass" : "",
  };
}

export const fieldsOf = (e: Entry): EntryFields => ({
  kind: e.kind, date: e.date, title: e.title, body: e.body, status: e.status,
  severity: e.severity, result: e.result, owner: e.owner, due: e.due,
});

/** Message for a form that cannot be saved yet, or "" when it can. */
export function entryProblem(f: EntryFields): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) return "Date must be yyyy-MM-dd.";
  if (!f.title.trim()) return "Title is required.";
  if (f.title.trim().length > 200) return "Title must be 200 characters or fewer.";
  if (f.due && !/^\d{4}-\d{2}-\d{2}$/.test(f.due)) return "Due must be yyyy-MM-dd.";
  return "";
}
