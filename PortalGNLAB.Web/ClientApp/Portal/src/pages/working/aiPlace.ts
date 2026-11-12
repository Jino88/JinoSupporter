// AI tab of a Working model: text typed or pasted is placed by AI (WorkingPlaceAiService) into Model Info fields, Dev Plan
// events and tasks, or new Dev History / Dev Issue / MP Issue / Review entries. The answer is shown item by item and only
// the items kept are saved, with the usual model and entry APIs.
import { issuesOfStage } from "./logic";
import { blankEvent, blankTask, readDevPlan, type PlanEvent } from "./plan";
import { phaseName, type Entry, type EntryFields, type EntryKind, type ModelInfo } from "./types";

export const INFO_KEYS = ["customer", "build", "phase", "owner", "line", "sopTarget", "note"] as const;
export type InfoKey = (typeof INFO_KEYS)[number];
export const INFO_LABEL: Record<InfoKey, string> = {
  customer: "Customer", build: "Build", phase: "Phase", owner: "PM owner", line: "Line", sopTarget: "SOP target", note: "Note",
};

export interface PlaceTask { id: string; name: string; start: string; end: string; owner: string; done: boolean }
export interface PlaceEvent { id: string; name: string; start: string; end: string; tasks: PlaceTask[] }
export interface PlaceEntry {
  kind: EntryKind;
  /** Issue: "dev" (Dev Issue) or "mp" (Mass Production Issues); "" for other kinds. */
  stage: "dev" | "mp" | "";
  date: string; title: string; body: string; owner: string; due: string;
  severity: string; result: string; phase: string; category: string; process: string; cause: string;
  /** Pictures sent with the text (0 = the first) that belong to this entry; saved with it. */
  images: number[];
}
export interface PlaceAnswer {
  info: Partial<Record<InfoKey, string>>;
  events: PlaceEvent[];
  entries: PlaceEntry[];
  note: string;
  model: string;
}

/** What the server is told about the model, so the answer can name events and tasks it already has and skip known entries. */
export function placeRequest(text: string, today: string, info: ModelInfo, entries: Entry[], images: { name: string; dataUrl: string }[] = []) {
  const plan = readDevPlan(info.devPlan);
  const mp = new Set(issuesOfStage(entries, "mp").map((e) => e.id));
  return {
    text, today, images,
    info: Object.fromEntries(INFO_KEYS.map((k) => [k, k === "phase" ? phaseName(info.phase) : info[k] ?? ""])) as Record<InfoKey, string>,
    events: plan.events.map((e) => ({
      id: e.id, name: e.name, start: e.start, end: e.end,
      tasks: e.tasks.map((t) => ({ id: t.id, name: t.name, start: t.start, end: t.end, owner: t.owner, done: t.done })),
    })),
    entries: entries.map((e) => ({
      id: e.id, kind: e.kind, stage: e.kind === "issue" ? (mp.has(e.id) ? "mp" : "dev") : "", seq: e.seq, date: e.date, title: e.title, status: e.status,
    })),
  };
}

const str = (x: unknown) => (typeof x === "string" ? x.trim() : "");
const list = (x: unknown) => (Array.isArray(x) ? x : []).filter((v): v is Record<string, unknown> => !!v && typeof v === "object");
const KINDS: EntryKind[] = ["dev", "issue", "review"];

/** The server answer read leniently. */
export function readPlaceAnswer(v: unknown): PlaceAnswer {
  const a = (v && typeof v === "object" ? v : {}) as Record<string, unknown>;
  const rawInfo = (a.info && typeof a.info === "object" ? a.info : {}) as Record<string, unknown>;
  const info: Partial<Record<InfoKey, string>> = {};
  for (const k of INFO_KEYS) if (typeof rawInfo[k] === "string") info[k] = str(rawInfo[k]);
  return {
    info,
    events: list(a.events).map((e) => ({
      id: str(e.id), name: str(e.name), start: str(e.start), end: str(e.end),
      tasks: list(e.tasks).map((t) => ({ id: str(t.id), name: str(t.name), start: str(t.start), end: str(t.end), owner: str(t.owner), done: t.done === true })),
    })),
    entries: list(a.entries).filter((e) => KINDS.includes(e.kind as EntryKind) && str(e.title)).map((e) => ({
      kind: e.kind as EntryKind,
      stage: e.kind !== "issue" ? "" : e.stage === "mp" || str(e.phase).toUpperCase() === "MP" ? "mp" : "dev",
      date: str(e.date), title: str(e.title), body: str(e.body), owner: str(e.owner), due: str(e.due),
      severity: str(e.severity), result: str(e.result), phase: str(e.phase), category: str(e.category), process: str(e.process), cause: str(e.cause),
      images: [...new Set((Array.isArray(e.images) ? e.images : []).filter((n): n is number => Number.isInteger(n) && n >= 0))],
    })),
    note: str(a.note),
    model: str(a.model),
  };
}

/** Where an item goes, in the order the places are shown. */
export const PLACES = ["Model Info", "Dev Plan", "Dev Issue", "Dev History", "Reviews", "MP Issues"] as const;
export type Place = (typeof PLACES)[number];
/** The tab of the model page each place is on. */
export const PLACE_TAB: Record<Place, "info" | "plan" | "devissue" | "dev" | "review" | "issue"> = {
  "Model Info": "info", "Dev Plan": "plan", "Dev Issue": "devissue", "Dev History": "dev", Reviews: "review", "MP Issues": "issue",
};

export const placeOf = (e: PlaceEntry): Place =>
  e.kind === "dev" ? "Dev History" : e.kind === "review" ? "Reviews" : e.stage === "mp" ? "MP Issues" : "Dev Issue";

/** A line of the answer that can be kept or left out (key), under the item it belongs to (parent: a task under its event). */
export interface PlaceItem {
  key: string;
  place: Place;
  /** New, or a change of something there already. */
  change: boolean;
  title: string;
  detail: string;
  parent?: string;
}

const days = (start: string, end: string) => (start ? (end && end !== start ? `${start} – ${end}` : start) : end);

/**
 * Every item of the answer to show and pick: Model Info fields that change, Dev Plan events (new, or an existing one whose
 * name or days change) and their tasks (new or changed), and the new entries. An existing event with nothing of its own
 * changing has no item; its tasks stand alone.
 */
export function placeItems(answer: PlaceAnswer, info: ModelInfo): PlaceItem[] {
  const out: PlaceItem[] = [];
  for (const k of INFO_KEYS) {
    const to = answer.info[k];
    if (to === undefined) continue;
    const from = (k === "phase" ? phaseName(info.phase) : info[k]) ?? "";
    if (to === from) continue;
    out.push({ key: `info.${k}`, place: "Model Info", change: !!from, title: INFO_LABEL[k], detail: from ? `${from} → ${to}` : to });
  }
  const plan = readDevPlan(info.devPlan);
  answer.events.forEach((e, i) => {
    const had = plan.events.find((x) => x.id && x.id === e.id);
    const evKey = `event.${i}`;
    let parent: string | undefined;
    if (!had) {
      if (!e.name) return;
      out.push({ key: evKey, place: "Dev Plan", change: false, title: e.name, detail: days(e.start, e.end) || "Days from its tasks" });
      parent = evKey;
    } else if ((e.name && e.name !== had.name) || (e.start && e.start !== had.start) || (e.end && e.end !== had.end)) {
      out.push({ key: evKey, place: "Dev Plan", change: true, title: e.name || had.name, detail: `${days(had.start, had.end) || "No date"} → ${days(e.start || had.start, e.end || had.end)}` });
      parent = evKey;
    }
    e.tasks.forEach((t, j) => {
      if (!t.name) return;
      const old = had?.tasks.find((x) => x.id && x.id === t.id);
      const text = [days(t.start, t.end), t.owner, t.done ? "done" : ""].filter(Boolean).join(" · ");
      out.push({
        key: `task.${i}.${j}`, place: "Dev Plan", change: !!old, parent,
        title: t.name, detail: `${had ? `in ${had.name} · ` : ""}${text || "No date"}`,
      });
    });
  });
  answer.entries.forEach((e, i) => {
    const pics = e.images.length ? `${e.images.length} picture${e.images.length > 1 ? "s" : ""}` : "";
    const bits = [e.date, e.owner, e.kind === "issue" ? [e.severity, e.due && `due ${e.due}`].filter(Boolean).join(" · ") : e.kind === "review" ? e.result : "", pics];
    out.push({ key: `entry.${i}`, place: placeOf(e), change: false, title: e.title, detail: bits.filter(Boolean).join(" · ") });
  });
  return out;
}

/**
 * The model information and new entries with the picked items put in. A task is put in only with its event when the event
 * is new (a new event left out takes its tasks with it); an existing event keeps everything not named in the answer.
 * Input: current info, the answer, the keys kept, today (date of an entry without one). Output: the info to save (or
 * null when it does not change), the entries to add and, for each, the pictures (by number) to save with it.
 */
export function applyPlace(info: ModelInfo, answer: PlaceAnswer, picked: Set<string>, today: string):
  { info: ModelInfo | null; entries: EntryFields[]; pictures: number[][] } {
  let next: ModelInfo = { ...info };
  let changed = false;
  for (const k of INFO_KEYS) {
    const to = answer.info[k];
    if (to !== undefined && picked.has(`info.${k}`) && to !== ((k === "phase" ? phaseName(info.phase) : info[k]) ?? "")) {
      next = { ...next, [k]: to };
      changed = true;
    }
  }

  const plan = readDevPlan(info.devPlan);
  let events: PlanEvent[] = plan.events;
  let planChanged = false;
  answer.events.forEach((e, i) => {
    const tasks = e.tasks.filter((t, j) => t.name && picked.has(`task.${i}.${j}`));
    const at = events.findIndex((x) => x.id && x.id === e.id);
    if (at < 0) {
      if (!e.name || !picked.has(`event.${i}`)) return;
      events = [...events, {
        ...blankEvent(events.length), name: e.name, start: e.start, end: e.end,
        tasks: tasks.map((t) => ({ ...blankTask(), name: t.name, start: t.start, end: t.end, owner: t.owner, done: t.done })),
      }];
      planChanged = true;
      return;
    }
    let ev = events[at];
    if (picked.has(`event.${i}`)) {
      ev = { ...ev, name: e.name || ev.name, start: e.start || ev.start, end: e.end || ev.end };
      planChanged = true;
    }
    for (const t of tasks) {
      const ti = ev.tasks.findIndex((x) => x.id && x.id === t.id);
      const fields = { name: t.name, start: t.start, end: t.end, owner: t.owner, done: t.done };
      ev = { ...ev, tasks: ti < 0 ? [...ev.tasks, { ...blankTask(), ...fields }] : ev.tasks.map((x, n) => (n === ti ? { ...x, ...fields } : x)) };
      planChanged = true;
    }
    events = events.map((x, n) => (n === at ? ev : x));
  });
  if (planChanged) { next = { ...next, devPlan: { ...plan, events } }; changed = true; }

  const kept = answer.entries.filter((_, i) => picked.has(`entry.${i}`));
  const entries = kept.map((e): EntryFields => {
    const issue = e.kind === "issue";
    return {
      kind: e.kind, date: /^\d{4}-\d{2}-\d{2}$/.test(e.date) ? e.date : today, title: e.title.slice(0, 200), body: e.body,
      status: issue ? "Open" : "", severity: issue ? e.severity : "", result: e.kind === "review" ? e.result || "Pass" : "",
      owner: e.owner, due: issue ? e.due : "", phase: issue ? (e.stage === "mp" ? "MP" : e.phase) : "",
      category: issue ? e.category : "", process: issue ? e.process : "", cause: issue ? e.cause : "", doneDate: "", finalResult: "",
    };
  });
  return { info: changed ? next : null, entries, pictures: kept.map((e) => e.images) };
}

/** Pictures AI can read: PNG, JPEG, GIF or WebP, 5 MB or less each, 6 at most with the ones already there. */
export const AI_PICTURE_TYPES = ["image/png", "image/jpeg", "image/gif", "image/webp"];
export const MAX_AI_PICTURES = 6;
export const MAX_AI_PICTURE_BYTES = 5 * 1024 * 1024;

/**
 * The pictures among pasted or dropped files (other files are ignored so a text paste still works), or a problem when
 * one cannot be sent. <have> = pictures already waiting.
 */
export function aiPictures<F extends { name: string; type: string; size: number }>(files: F[], have: number): { files: F[]; problem: string } {
  const pics = files.filter((f) => f.type.startsWith("image/"));
  for (const f of pics) {
    if (!AI_PICTURE_TYPES.includes(f.type)) return { files: [], problem: `${f.name || "This picture"} is not PNG, JPEG, GIF or WebP.` };
    if (f.size > MAX_AI_PICTURE_BYTES) return { files: [], problem: `${f.name || "This picture"} is larger than 5 MB.` };
  }
  if (have + pics.length > MAX_AI_PICTURES) return { files: [], problem: `Up to ${MAX_AI_PICTURES} pictures can be sent.` };
  return { files: pics, problem: "" };
}
