import {
  IMAGE_TYPES, MAX_IMAGE_BYTES, MAX_IMAGES_PER_UPLOAD, MAX_UPLOAD_BYTES, MAX_VIDEO_BYTES, phaseName, VIDEO_EXTENSIONS, VIDEO_TYPES,
} from "./types";
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
  phase: phaseName(e.phase), category: e.category ?? "", process: e.process ?? "", cause: e.cause ?? "",
  doneDate: e.doneDate ?? "", finalResult: e.finalResult ?? "",
});

/** Issue fields with Status kept in step with the done date. */
export const issueFields = (f: EntryFields): EntryFields => ({ ...f, status: f.doneDate ? "Closed" : "Open" });

/**
 * The place (1, 2 …) of each top issue among the top issues in number order: the No. of a list numbered in order (My Plans),
 * where the register numbers have gaps (sub-issues and removed issues take numbers too).
 */
export function orderNumbers(list: Entry[]): Map<string, number> {
  return new Map(issueTree(list).top.map((e, i) => [e.id, i + 1]));
}

/** Issues in number order (No. 1 first). */
export const issuesBySeq = (list: Entry[]) => list.filter((e) => e.kind === "issue").sort((a, b) => a.seq - b.seq || a.createdAt.localeCompare(b.createdAt));

/**
 * Issues as a two-level list: the top issues in number order and the sub-issues of each, also in number order.
 * A sub-issue whose issue is not in the list (e.g. another setup row) is shown as a top issue.
 */
export function issueTree(list: Entry[]): { top: Entry[]; subs: Map<string, Entry[]> } {
  const issues = issuesBySeq(list);
  const ids = new Set(issues.map((e) => e.id));
  const subs = new Map<string, Entry[]>();
  const top: Entry[] = [];
  for (const e of issues) {
    if (e.parentId && e.parentId !== e.id && ids.has(e.parentId)) subs.set(e.parentId, [...(subs.get(e.parentId) ?? []), e]);
    else top.push(e);
  }
  return { top, subs };
}

export type IssueStage = "dev" | "mp";

/**
 * Where an issue is listed: Mass Production when it was raised in MP or against a Setup Plan row, Development otherwise.
 * A sub-issue is listed with its issue.
 */
export function issuesOfStage(list: Entry[], stage: IssueStage): Entry[] {
  const byId = new Map(list.map((e) => [e.id, e]));
  const own = (e: Entry): IssueStage => (e.phase === "MP" || e.setupRowId ? "mp" : "dev");
  return list.filter((e) => {
    if (e.kind !== "issue") return false;
    const parent = e.parentId && e.parentId !== e.id ? byId.get(e.parentId) : undefined;
    return own(parent ?? e) === stage;
  });
}

/** All entries with the ones a part of the page was given (some) swapped for what it gave back (list). */
export function replaceSome(all: Entry[], some: Entry[], list: Entry[]): Entry[] {
  const mine = new Set(some.map((e) => e.id));
  return [...all.filter((e) => !mine.has(e.id)), ...list];
}

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

export const blankLog = (today: string): LogFields => ({ date: today, text: "", detail: "", owner: "", due: "", doneDate: "" });

/** A note's one-line issue and its detail; lines after the first of an older several-line text go to the front of the detail. */
export function noteParts(l: { text: string; detail?: string }): { line: string; detail: string } {
  const [line = "", ...rest] = l.text.replace(/\r\n?/g, "\n").trim().split("\n");
  return { line: line.trim(), detail: [rest.join("\n").trim(), (l.detail ?? "").trim()].filter(Boolean).join("\n") };
}

export const logFieldsOf = (l: EntryLog): LogFields => {
  const { line, detail } = noteParts(l);
  return { date: l.date, text: line, detail, owner: l.owner ?? "", due: l.due ?? "", doneDate: l.doneDate ?? "" };
};

/** Notes in date order, oldest first. */
export const notesInOrder = (logs: EntryLog[]) =>
  [...logs].sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));

/** A note that is not done and whose due date has passed. */
export const isLogLate = (l: LogFields, today: string) => !l.doneDate && !!l.due && l.due < today;

export function logProblem(f: LogFields): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) return "Date must be a date (yyMMdd).";
  if (!f.text.trim()) return "Write what was done.";
  if (/[\r\n]/.test(f.text)) return "Write the issue in one line and the rest in Issue detail.";
  if (f.text.trim().length > 500) return "The issue must be 500 characters or fewer.";
  if (f.detail.trim().length > 5000) return "Issue detail must be 5000 characters or fewer.";
  if (f.owner.trim().length > 100) return "Owner must be 100 characters or fewer.";
  if (f.due && !/^\d{4}-\d{2}-\d{2}$/.test(f.due)) return "Due must be a date (yyMMdd).";
  if (f.doneDate && !/^\d{4}-\d{2}-\d{2}$/.test(f.doneDate)) return "Done date must be a date (yyMMdd).";
  if (f.doneDate && f.doneDate < f.date) return "Done date cannot be before the note date.";
  return "";
}

type FileLike = { name: string; type: string; size: number };

/** "image" / "video" for a picture or video (a video with no type is told by its extension), otherwise null. */
export function mediaKind(f: FileLike): "image" | "video" | null {
  if (f.type.startsWith("image/")) return "image";
  if (f.type.startsWith("video/")) return "video";
  const ext = f.name.split(".").pop()?.toLowerCase() ?? "";
  return (!f.type || f.type === "application/octet-stream") && VIDEO_EXTENSIONS[ext] ? "video" : null;
}

/**
 * Pictures and videos among pasted, dropped or picked files. Other files are ignored so a normal text paste still works;
 * a file the server would refuse comes back as a problem instead. <have> / <haveBytes> are the files already waiting.
 */
export function pastedImages<F extends FileLike>(files: F[], have: number, haveBytes = 0): { files: F[]; problem: string } {
  const media = files.filter((f) => mediaKind(f));
  const label = (f: F) => f.name || "This file";
  for (const f of media) {
    const ext = f.name.split(".").pop()?.toLowerCase() ?? "";
    if (mediaKind(f) === "image") {
      if (!IMAGE_TYPES.includes(f.type)) return { files: [], problem: `${label(f)} is not PNG, JPEG, GIF, WebP or BMP.` };
      if (f.size > MAX_IMAGE_BYTES) return { files: [], problem: `${label(f)} is larger than 10 MB.` };
    } else {
      if (!VIDEO_TYPES.includes(f.type) && !VIDEO_EXTENSIONS[ext])
        return { files: [], problem: `${label(f)} is not an MP4, MOV, WebM, MKV or AVI video.` };
      if (f.size > MAX_VIDEO_BYTES) return { files: [], problem: `${label(f)} is larger than 500 MB.` };
    }
  }
  if (have + media.length > MAX_IMAGES_PER_UPLOAD) return { files: [], problem: `Add at most ${MAX_IMAGES_PER_UPLOAD} files at a time.` };
  if (haveBytes + media.reduce((n, f) => n + f.size, 0) > MAX_UPLOAD_BYTES)
    return { files: [], problem: "The files of one save must be 1 GB or less in total." };
  return { files: media, problem: "" };
}

/** Message for a form that cannot be saved yet, or "" when it can. */
export function entryProblem(f: EntryFields): string {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(f.date)) return "Date must be a date (yyMMdd).";
  if (!f.title.trim()) return "Title is required.";
  if (f.title.trim().length > 200) return "Title must be 200 characters or fewer.";
  if (f.body.length > 20000) return "Issue detail must be 20000 characters or fewer.";
  if (f.due && !/^\d{4}-\d{2}-\d{2}$/.test(f.due)) return "Due must be a date (yyMMdd).";
  if (f.doneDate && !/^\d{4}-\d{2}-\d{2}$/.test(f.doneDate)) return "Done date must be a date (yyMMdd).";
  if (f.doneDate && f.doneDate < f.date) return "Done date cannot be before the registered date.";
  return "";
}
