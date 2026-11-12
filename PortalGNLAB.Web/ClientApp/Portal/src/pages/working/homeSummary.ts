// Home of Working: a summary of every model, the development plans of the models in development first, then the MP plans
// (the Setup Plan of each model).
import { addDays } from "../productionInput/logic";
import { isLate, isOpenIssue, issuesOfStage, type GroupNode } from "./logic";
import { COLORS, eventSpan, progress, readDevPlan, taskSpan } from "./plan";
import { rowsInOrder, type SetupRow } from "./setup";
import { phaseName, type Entry, type ModelInfo, type WorkingModel } from "./types";

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
}

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
      const s = taskSpan(t);
      return { id: t.id, name: t.name || "(no name)", span: s, done: t.done, late: !t.done && !!s && s[1] < today, owner: t.owner };
    });
    const lateTasks = tasks.filter((t) => t.late).length;
    return { id: ev.id, name: ev.name || "(no name)", color: ev.color, span, done, total, lateTasks, state: eventState(span, done, total, today), tasks };
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
  const lines: string[] = [];
  for (const r of rowsInOrder(rows)) { const l = r.line.trim() || "(No line)"; if (!lines.includes(l)) lines.push(l); }
  const events = lines.map((line, i) => ({
    id: "line:" + line, name: line, start: "", end: "", color: COLORS[(i + 1) % COLORS.length], rows: [], bomFrom: "",
    tasks: rowsInOrder(rows).filter((r) => (r.line.trim() || "(No line)") === line).map((r) => ({
      id: r.id, name: r.processName || `No. ${r.no}`, owner: "", note: "",
      start: r.machinePlan || r.setupPlan, end: r.setupPlan || r.machinePlan, done: r.setupStatus === "O" || !!r.setupFinish,
    })),
  }));
  return planBrief(name, path, { phase, devPlan: { events } }, today);
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
  const spans = plans.flatMap((p) => p.events.flatMap((e) => [e.span, ...e.tasks.map((t) => t.span)])).filter((s): s is [string, string] => !!s);
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

/** An MP issue of a model for the MP Issues list. */
export interface MpIssueRow {
  model: string;
  path: string[];
  issue: Entry;
  open: boolean;
  late: boolean;
}

/**
 * The Mass Production issues of every model (raised in MP or on a Setup Plan row, with their sub-issues): open ones first,
 * late ones first among them, then in the order of the models and by number.
 */
export function mpIssueRows(loads: ModelLoad[], today: string): MpIssueRow[] {
  const rows = loads.flatMap((l, at) => issuesOfStage(l.entries, "mp").map((issue) => ({
    model: l.name, path: l.path, issue, open: isOpenIssue(issue), late: isLate(issue, today), at,
  })));
  rows.sort((a, b) => Number(!a.open) - Number(!b.open) || Number(!a.late) - Number(!b.late) || a.at - b.at || a.issue.seq - b.issue.seq);
  return rows.map(({ at: _at, ...r }) => r);
}
