// Names of the Dev Plan made of the standard names (planVocab): an event = its stage and quantity, a task = its standard task
// and what is particular to it. Also the tidy-up of old plans (written names matched to standard ones) and the tasks an
// event of a stage starts with.
import { blankTask, type DevPlan, type PlanEvent, type PlanTask } from "./plan";
import type { PlanVocab, VocabItem } from "./planVocab";

/** "ER (1.2k)"; the stage alone without a quantity. */
export const eventLabel = (stage: string, qty: string) => `${stage}${qty.trim() ? ` (${qty.trim()})` : ""}`;
/** "Sample shipment · 300ea Bringup"; the standard name alone without a detail. */
export const taskLabel = (name: string, detail: string) => `${name}${detail.trim() ? ` · ${detail.trim()}` : ""}`;

export const itemOf = (code: string | undefined, items: VocabItem[]) => (code ? items.find((i) => i.code === code) : undefined);

/** What is left of a name once the standard part is taken out, without the separators and brackets around it. */
function leftOver(name: string, at: number, length: number): string {
  let rest = (name.slice(0, at) + " " + name.slice(at + length)).replace(/\s+/g, " ");
  for (let before = ""; before !== rest;) {
    before = rest;
    rest = rest.trim().replace(/^[:\-_/·,.|~]+|[:\-_/·,.|~]+$/g, "").trim();
    if (/^\([^()]*\)$/.test(rest) || /^\[[^[\]]*\]$/.test(rest)) rest = rest.slice(1, -1);
  }
  return rest.replace(/\(\s*\)/g, "").replace(/\s+([:,])/g, "$1").trim();
}

/** The standard task a written name is: its name, code or the longest of its other names found in it; null if none. */
export function matchTask(name: string, items: VocabItem[]): { item: VocabItem; detail: string } | null {
  const low = name.toLowerCase();
  let best: { item: VocabItem; at: number; length: number } | null = null;
  for (const item of items) {
    for (const word of [item.name, item.code, ...item.aliases]) {
      const w = word.trim().toLowerCase();
      if (w.length < 2) continue;
      const at = low.indexOf(w);
      if (at >= 0 && (!best || w.length > best.length)) best = { item, at, length: w.length };
    }
  }
  return best && { item: best.item, detail: leftOver(name, best.at, best.length) };
}

/** The standard stage a written event name starts with ("ER (1.2kpcs)" = ER, 1.2kpcs); null if none. */
export function matchStage(name: string, stages: VocabItem[]): { item: VocabItem; qty: string } | null {
  const esc = (x: string) => x.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  let best: { item: VocabItem; length: number } | null = null;
  for (const item of stages) {
    for (const word of [item.code, item.name, ...item.aliases]) {
      const m = word.trim() && new RegExp(`^\\s*${esc(word.trim())}(?![A-Za-z0-9])`, "i").exec(name);
      if (m && (!best || m[0].length > best.length)) best = { item, length: m[0].length };
    }
  }
  return best && { item: best.item, qty: leftOver(name, 0, best.length) };
}

/** A name of an old plan matched to a standard one, to be looked at and applied. */
export interface TidyRow {
  key: string;
  kind: "event" | "task";
  eventId: string;
  taskId?: string;
  /** The name as written. */
  from: string;
  /** Stage or task code chosen ("" = leave it as written). */
  code: string;
  /** Quantity of an event or detail of a task. */
  detail: string;
}

/** Every event and task of the plan without a standard name yet, with the standard one its name looks like (or ""). */
export function tidyRows(plan: DevPlan, vocab: PlanVocab): TidyRow[] {
  return plan.events.flatMap((e) => {
    const out: TidyRow[] = [];
    if (!e.stage) {
      const m = matchStage(e.name, vocab.stages);
      out.push({ key: e.id, kind: "event", eventId: e.id, from: e.name, code: m?.item.code ?? "", detail: m?.qty ?? "" });
    }
    for (const t of e.tasks.filter((t) => !t.code)) {
      const m = matchTask(t.name, vocab.tasks);
      out.push({ key: `${e.id}/${t.id}`, kind: "task", eventId: e.id, taskId: t.id, from: t.name, code: m?.item.code ?? "", detail: m?.detail ?? t.name });
    }
    return out;
  });
}

/** An event given a stage (and quantity, kept as typed; trimmed when kept), its name made of them. */
export const withStage = (e: PlanEvent, stage: VocabItem, qty: string): PlanEvent =>
  ({ ...e, stage: stage.code, qty: qty || undefined, name: eventLabel(stage.name, qty) });
/** A task given a standard task (and detail, kept as typed; trimmed when kept), its name made of them. */
export const withCode = (t: PlanTask, task: VocabItem, detail: string): PlanTask =>
  ({ ...t, code: task.code, detail: detail || undefined, name: taskLabel(task.name, detail) });
/** An event or task with no standard name any more: the name stays as it was, the code goes. */
export const freeEvent = (e: PlanEvent): PlanEvent => ({ ...e, stage: undefined, qty: undefined });
export const freeTask = (t: PlanTask): PlanTask => ({ ...t, code: undefined, detail: undefined });

/** The plan with the rows given a code renamed; rows left without a code, and names not in the rows, stay as they are. */
export function applyTidy(plan: DevPlan, rows: TidyRow[], vocab: PlanVocab): DevPlan {
  const byKey = new Map(rows.filter((r) => r.code).map((r) => [r.key, r]));
  return {
    ...plan,
    events: plan.events.map((e) => {
      const er = byKey.get(e.id), stage = itemOf(er?.code, vocab.stages);
      const ev = er && stage ? withStage(e, stage, er.detail.trim()) : e;
      return {
        ...ev,
        tasks: ev.tasks.map((t) => {
          const tr = byKey.get(`${e.id}/${t.id}`), task = itemOf(tr?.code, vocab.tasks);
          return tr && task ? withCode(t, task, tr.detail.trim()) : t;
        }),
      };
    }),
  };
}

/** The standard tasks of the event's stage it does not have yet, in the template's order, as new tasks. */
export function templateTasks(ev: PlanEvent, vocab: PlanVocab): PlanTask[] {
  const have = new Set(ev.tasks.map((t) => t.code).filter(Boolean));
  return (ev.stage ? vocab.templates[ev.stage] ?? [] : [])
    .filter((c) => !have.has(c))
    .map((c) => itemOf(c, vocab.tasks))
    .filter((i): i is VocabItem => !!i)
    .map((i) => withCode(blankTask(), i, ""));
}
