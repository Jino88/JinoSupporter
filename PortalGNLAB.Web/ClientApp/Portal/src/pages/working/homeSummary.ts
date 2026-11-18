// Home of Working: a summary of every model, the development plans of the models in development first, then the MP plans
// (the Setup Plan of each model).
import { addDays } from "./dates";
import { isLate, isLogLate, isOpenIssue, issuesOfStage, issueTree, notesInOrder, type GroupNode, type IssueStage } from "./logic";
import { actualSpan, eventSpan, groupSpan, progress, readDevPlan } from "./plan";
import { setupEvents, type SetupRow } from "./setup";
import { phaseName, type Entry, type EntryLog, type ModelInfo, type WorkingModel } from "./types";

/** A model of the tree as Home loaded it: its information and entries, or the error it could not be loaded with. */
export interface ModelLoad {
  name: string;
  path: string[];
  model: WorkingModel | null;
  entries: Entry[];
  /** Setup Plan rows (Mass Production); none when the model has no Setup Plan. */
  setup?: SetupRow[];
  error: string;
}

/** Every model of the tree once, in the order the tree shows them (a group's sub-groups before its own models), with its path. */
export function treeModels(nodes: GroupNode[], trail: string[] = [], seen = new Set<string>()): { name: string; path: string[] }[] {
  return nodes.flatMap((n) => {
    const path = [...trail, n.name];
    const below = treeModels(n.children, path, seen);
    const own = n.models.filter((m) => !seen.has(m)).map((m) => { seen.add(m); return { name: m, path }; });
    return [...below, ...own];
  });
}

/** yyyy-MM-dd as yyyy/MM/dd. */
export const slashDate = (d: string) => d.replace(/-/g, "/");
/** yy/MM/dd, e.g. 26/01/03 (MP Issues). */
export const shortDate = (d: string) => slashDate(d).slice(2);

/** Where an event is against today: finished, under way, coming, past its end with tasks open, or without dates. */
export type EventState = "done" | "now" | "upcoming" | "late" | "undated";

/** A task of an event: its days, done, late (not done and its last day before today) and owner. */
export interface TaskBrief {
  id: string;
  name: string;
  span: [string, string] | null;
  done: boolean;
  late: boolean;
  owner: string;
  /** Of a sub-task: the id of its task, shown (and folded away) under it. */
  parent?: string;
  /** The days it really took (actual start to actual end, or to today while still going); absent before it starts. */
  actual?: Actual;
}

/** What really happened against the plan: the days, ended (or still going) after the planned end, not ended yet. */
export interface Actual {
  span: [string, string];
  over: boolean;
  going: boolean;
}

const actualOf = (span: [string, string] | null, plan: [string, string] | null, going: boolean): Actual | undefined =>
  span ? { span, over: !!plan && span[1] > plan[1], going } : undefined;

export interface EventBrief {
  id: string;
  name: string;
  color: string;
  span: [string, string] | null;
  done: number;
  total: number;
  /** Tasks not done whose end (or start) is before today. */
  lateTasks: number;
  state: EventState;
  /** Every task of the event, in the order of the plan. */
  tasks: TaskBrief[];
  /** From the first actual start of its tasks to the last actual end (today while one is still going). */
  actual?: Actual;
}

export interface PlanBrief {
  name: string;
  path: string[];
  phase: string;
  events: EventBrief[];
  /** Events under way today. */
  now: EventBrief[];
  /** First event still to come. */
  next: EventBrief | null;
  lateTasks: number;
}

function eventState(span: [string, string] | null, done: number, total: number, today: string): EventState {
  if (total > 0 && done === total) return "done";
  if (!span) return "undated";
  if (span[1] < today) return total > 0 ? "late" : "done";
  return span[0] <= today ? "now" : "upcoming";
}

/** The development plan of a model in a few lines: each event with its days, tasks done and state, the one under way and the next. */
export function planBrief(name: string, path: string[], info: ModelInfo, today: string): PlanBrief {
  const events = readDevPlan(info.devPlan).events.map((ev): EventBrief => {
    const span = eventSpan(ev);
    const [done, total] = progress(ev);
    const tasks = ev.tasks.map((t): TaskBrief => {
      const s = groupSpan(ev.tasks, t);
      const kids = ev.tasks.filter((c) => c.parentId === t.id);
      const going = [t, ...kids].some((x) => !!x.actualStart && !x.actualEnd && !x.done);
      const actual = actualOf(actualSpan(ev.tasks, t, today), s, going);
      return {
        id: t.id, name: t.name || "(no name)", span: s, done: t.done, late: !t.done && !!s && s[1] < today, owner: t.owner,
        ...(t.parentId ? { parent: t.parentId } : {}), ...(actual ? { actual } : {}),
      };
    });
    const lateTasks = tasks.filter((t) => t.late).length;
    const acts = tasks.filter((t) => !t.parent && t.actual).map((t) => t.actual!);
    const actual = acts.length
      ? actualOf([acts.map((a) => a.span[0]).sort()[0], acts.map((a) => a.span[1]).sort().at(-1)!], span, acts.some((a) => a.going))
      : undefined;
    return {
      id: ev.id, name: ev.name || "(no name)", color: ev.color, span, done, total, lateTasks, state: eventState(span, done, total, today), tasks,
      ...(actual ? { actual } : {}),
    };
  });
  const upcoming = events.filter((e) => e.state === "upcoming").sort((a, b) => a.span![0].localeCompare(b.span![0]));
  return {
    name, path, phase: phaseName(info.phase), events,
    now: events.filter((e) => e.state === "now"),
    next: upcoming[0] ?? null,
    lateTasks: events.reduce((n, e) => n + e.lateTasks, 0),
  };
}

/**
 * The plans of the models in development (any phase but MP) that have a plan: the ones with an event under way first,
 * then by the start of their next event, then by name; and the models in a development phase (ER to PV) with no plan yet.
 */
export function devPlans(loads: ModelLoad[], today: string): { plans: PlanBrief[]; noPlan: string[] } {
  const dev = loads.filter((l) => l.model && phaseName(l.model.info.phase) !== "MP");
  const plans = dev.map((l) => planBrief(l.name, l.path, l.model!.info, today)).filter((p) => p.events.length);
  const firstDay = (p: PlanBrief) => p.next?.span?.[0] ?? "9999";
  plans.sort((a, b) => Number(!a.now.length) - Number(!b.now.length) || firstDay(a).localeCompare(firstDay(b)) || a.name.localeCompare(b.name));
  const has = new Set(plans.map((p) => p.name));
  const noPlan = dev.filter((l) => !has.has(l.name) && !!phaseName(l.model!.info.phase)).map((l) => l.name);
  return { plans, noPlan };
}

/**
 * The Setup Plan of a model as a plan: a line per event and its processes as tasks, each from its machine plan day to its
 * setup plan day (one of them when only one is set), done when its setup is OK or has a finish day.
 */
export function setupPlanBrief(name: string, path: string[], phase: string, rows: SetupRow[], today: string): PlanBrief {
  return planBrief(name, path, { phase, devPlan: { events: setupEvents(rows) } }, today);
}

/** The MP plans: every model with a Setup Plan, in the same order as the development plans. */
export function mpPlans(loads: ModelLoad[], today: string): PlanBrief[] {
  const plans = loads.filter((l) => l.model && l.setup?.length)
    .map((l) => setupPlanBrief(l.name, l.path, phaseName(l.model!.info.phase), l.setup!, today));
  const firstDay = (p: PlanBrief) => p.next?.span?.[0] ?? "9999";
  return plans.sort((a, b) => Number(!a.now.length) - Number(!b.now.length) || firstDay(a).localeCompare(firstDay(b)) || a.name.localeCompare(b.name));
}

/**
 * The days the plan timeline shows: from the first day of the plans (events and tasks) to the last, today always in, a few days on each side.
 * Output: first and last day and the number of days, or null without any dated event.
 */
export function timelineRange(plans: PlanBrief[], today: string, pad = 3): { from: string; to: string; days: number } | null {
  const spans = plans.flatMap((p) => p.events.flatMap((e) => [e.span, e.actual?.span, ...e.tasks.flatMap((t) => [t.span, t.actual?.span])]))
    .filter((s): s is [string, string] => !!s);
  if (!spans.length) return null;
  const from = addDays([today, ...spans.map((s) => s[0])].sort()[0], -pad);
  const to = addDays([today, ...spans.map((s) => s[1])].sort().at(-1)!, pad);
  return { from, to, days: dayCount(from, to) + 1 };
}

/** Days from one yyyy-MM-dd to another. */
export const dayCount = (from: string, to: string) =>
  Math.round((Date.parse(to + "T00:00:00Z") - Date.parse(from + "T00:00:00Z")) / 86400000);

/** Left and width (in % of the timeline) of a span of days. */
export function barAt(span: [string, string], range: { from: string; days: number }): { left: number; width: number } {
  const left = (dayCount(range.from, span[0]) / range.days) * 100;
  const width = ((dayCount(span[0], span[1]) + 1) / range.days) * 100;
  return { left, width };
}

/**
 * Where the days of a bar are written: beside it, on its right when there is room there (<need> % of the timeline), else on
 * its left when there is room there, else on the side with more room; only a bar leaving almost no room on either side
 * (under 8 %) has them inside. Every bar alike, so none is cut to "2026/1…".
 */
export function labelSide(bar: { left: number; width: number }, need = 14): "in" | "right" | "left" {
  const right = 100 - bar.left - bar.width, left = bar.left;
  if (right >= need) return "right";
  if (left >= need) return "left";
  if (Math.max(left, right) < 8) return "in";
  return right >= left ? "right" : "left";
}

/** The months over the timeline: [yyyy-MM, left %, width %], for its header. */
export function timelineMonths(range: { from: string; to: string; days: number }): [string, number, number][] {
  const out: [string, number, number][] = [];
  let d = range.from;
  while (d <= range.to) {
    const m = d.slice(0, 7);
    const next = addDays(`${m}-01`, 32).slice(0, 7) + "-01";
    const end = addDays(next, -1) < range.to ? addDays(next, -1) : range.to;
    const { left, width } = barAt([d, end], range);
    out.push([m, left, width]);
    d = next;
  }
  return out;
}

export interface ModelRow {
  name: string;
  path: string[];
  phase: string;
  customer: string;
  build: string;
  owner: string;
  sopTarget: string;
  events: number;
  devOpen: number;
  mpOpen: number;
  lateIssues: number;
  updatedAt: string;
  error: string;
}

/** A line of the model table of Home: the model information, its plan events and its open issues (development, mass production, late). */
export function modelRow(l: ModelLoad, today: string): ModelRow {
  const info = l.model?.info ?? {};
  const open = (list: Entry[]) => list.filter(isOpenIssue).length;
  return {
    name: l.name, path: l.path, phase: phaseName(info.phase),
    customer: info.customer ?? "", build: info.build ?? "", owner: info.owner ?? "", sopTarget: info.sopTarget ?? "",
    events: readDevPlan(info.devPlan).events.length,
    devOpen: open(issuesOfStage(l.entries, "dev")), mpOpen: open(issuesOfStage(l.entries, "mp")),
    lateIssues: l.entries.filter((e) => isLate(e, today)).length,
    updatedAt: l.model?.updatedAt ?? "", error: l.error,
  };
}

/** A note of an issue that is not done, numbered as the Issues tab numbers it (issue No.-note No.). */
export interface OpenNote {
  no: string;
  log: EntryLog;
  late: boolean;
}

/** An issue of a model for the Dev Issues or MP Issues list, with its sub-issues (none for a sub-issue) and the notes not done. */
export interface IssueRow {
  model: string;
  path: string[];
  /** No. as the Issues tab shows it: the number, or issue No.sub-issue No. for a sub-issue. */
  no: string;
  issue: Entry;
  open: boolean;
  late: boolean;
  subs: IssueRow[];
  openNotes: OpenNote[];
}

/** Whether an issue or one of its sub-issues is open (an issue stays in the open list while a sub-issue is open). */
export const hasOpen = (r: IssueRow) => r.open || r.subs.some((s) => s.open);

/**
 * The issues of a stage of every model (Development, or Mass Production: raised in MP or on a Setup Plan row), each with its
 * sub-issues and the notes not done: issues with something open first, late ones first among them, then in the order of
 * the models and by number.
 */
export function issueRows(loads: ModelLoad[], today: string, stage: IssueStage): IssueRow[] {
  const row = (l: ModelLoad, issue: Entry, no: string, subs: IssueRow[]): IssueRow => ({
    model: l.name, path: l.path, no, issue, open: isOpenIssue(issue), late: isLate(issue, today), subs,
    openNotes: notesInOrder(issue.logs ?? []).flatMap((log, i) => (log.doneDate ? [] : [{ no: `${no}-${i + 1}`, log, late: isLogLate(log, today) }])),
  });
  const rows = loads.flatMap((l, at) => {
    const { top, subs } = issueTree(issuesOfStage(l.entries, stage));
    return top.map((issue) => {
      const no = String(issue.seq || "—");
      return { ...row(l, issue, no, (subs.get(issue.id) ?? []).map((c, i) => row(l, c, `${no}.${i + 1}`, []))), at };
    });
  });
  const late = (r: IssueRow) => r.late || r.subs.some((s) => s.late);
  rows.sort((a, b) => Number(!hasOpen(a)) - Number(!hasOpen(b)) || Number(!late(a)) - Number(!late(b)) || a.at - b.at || a.issue.seq - b.issue.seq);
  return rows.map(({ at: _at, ...r }) => r);
}
