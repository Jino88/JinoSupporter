import { userName } from "../../userNames";
import { Fragment, useEffect, useState, type CSSProperties, type ReactNode } from "react";
import { workingApi } from "./api";
import { useTr } from "./i18n";
import { setupApi } from "./setup";
import {
  barAt, devPlans, hasOpen, issueRows, labelSide, modelRow, mpPlans, shortDate, slashDate, timelineMonths, timelineRange,
  type Actual, type EventBrief, type IssueRow, type ModelLoad, type PlanBrief, type TaskBrief,
} from "./homeSummary";
import { noteParts, type IssueStage } from "./logic";

const STATE_LABEL: Record<EventBrief["state"], string> = { done: "Done", now: "Now", upcoming: "Next", late: "Late", undated: "No date" };
/** Models loaded at the same time, so a long tree does not send every request at once. */
const AT_ONCE = 6;

/**
 * The events whose tasks are folded away ("kind|model|event") and the tasks whose sub-tasks are ("kind|model|event|task"),
 * kept in the browser (all unfolded at first).
 */
const FOLD_KEY = "working.home.folded";
const savedFolds = (): string[] => {
  try {
    const v = JSON.parse(localStorage.getItem(FOLD_KEY) ?? "[]");
    return Array.isArray(v) ? v.filter((k): k is string => typeof k === "string") : [];
  } catch { return []; }
};
const saveFolds = (keys: Set<string>) => { try { localStorage.setItem(FOLD_KEY, JSON.stringify([...keys])); } catch { /* private mode */ } };

const spanText = (e: EventBrief) => (e.span ? `${slashDate(e.span[0])}${e.span[1] !== e.span[0] ? ` – ${slashDate(e.span[1]).slice(5)}` : ""}` : "No date");

/** Every model's information and entries, a few at a time; a model that fails keeps its error and the others go on. */
async function loadAll(models: { name: string; path: string[] }[], onCount: (n: number) => void): Promise<ModelLoad[]> {
  const out: ModelLoad[] = new Array(models.length);
  let at = 0, done = 0;
  const worker = async () => {
    while (at < models.length) {
      const i = at++;
      const { name, path } = models[i];
      try {
        const [model, entries, setup] = await Promise.all([
          workingApi.model(name), workingApi.entries(name), setupApi.rows(name).catch(() => []),
        ]);
        out[i] = { name, path, model, entries, setup, error: "" };
      } catch (e) {
        out[i] = { name, path, model: null, entries: [], error: (e as Error).message };
      }
      onCount(++done);
    }
  };
  await Promise.all(Array.from({ length: Math.min(AT_ONCE, models.length) }, worker));
  return out;
}

/** Home of Working: the development plans of the models in development, then every model in a table. */
export default function Home(props: {
  models: { name: string; path: string[] }[] | null;
  today: string;
  onOpen: (model: string, tab: "info" | "plan" | "setup" | "devissue" | "issue") => void;
  /** Home itself, or the Dev issues or MP issues of every model. */
  view?: "home" | "devissues" | "mpissues";
  /** Already loaded (tests and the page reopening); loaded here when left out. */
  initial?: ModelLoad[];
  /** The events ("kind|model|event") and tasks ("…|task") folded at first (tests); else the ones folded last time in this browser. */
  folded?: string[];
}) {
  const tr = useTr();
  const [loads, setLoads] = useState<ModelLoad[] | null>(props.initial ?? null);
  const [count, setCount] = useState(0);
  const key = props.models?.map((m) => m.name).join("\n");

  useEffect(() => {
    if (props.initial || !props.models) return;
    let live = true;
    setLoads(null);
    setCount(0);
    loadAll(props.models, (n) => live && setCount(n)).then((l) => live && setLoads(l));
    return () => { live = false; };
  }, [key]);

  if (!loads) {
    return <div className="card pi-empty">{props.models ? `Loading models… ${count} / ${props.models.length}` : "Loading…"}</div>;
  }
  if (props.view === "devissues" || props.view === "mpissues") {
    return <StageIssues key={props.view} stage={props.view === "devissues" ? "dev" : "mp"} loads={loads} today={props.today} onOpen={props.onOpen} />;
  }
  const { plans, noPlan } = devPlans(loads, props.today);
  const mp = mpPlans(loads, props.today);
  const rows = loads.map((l) => modelRow(l, props.today));
  const failed = rows.filter((r) => r.error);

  return (
    <div className="wk-home">
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>{tr("Development plans")}<span className="cnt">{plans.length}</span></h2>
          <div className="pi-legend">
            <span>{tr("Models in development (not MP) with a Dev Plan")} · {tr("today")} {slashDate(props.today)}</span>
          </div>
        </div>
        <div className="pi-sec-b">
          {plans.length ? <PlanTimeline kind="dev" plans={plans} today={props.today} onOpen={props.onOpen} folded={props.folded} /> : (
            <div className="muted small">{tr("No model in development has a Dev Plan yet. Add events in a model's Development › Dev Plan.")}</div>
          )}
          {noPlan.length > 0 && (
            <div className="wk-hm-noplan">
              <span className="muted small">{tr("In development without a plan:")}</span>
              {noPlan.map((m) => <button key={m} className="wk-chip" onClick={() => props.onOpen(m, "plan")}>{m}</button>)}
            </div>
          )}
        </div>
      </section>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>{tr("MP plans")}<span className="cnt">{mp.length}</span></h2>
          <div className="pi-legend">
            <span>{tr("Plan of each model: each process from its machine plan to its setup plan, by line")} · {tr("today")} {slashDate(props.today)}</span>
          </div>
        </div>
        <div className="pi-sec-b">
          {mp.length ? <PlanTimeline kind="mp" plans={mp} today={props.today} onOpen={props.onOpen} folded={props.folded} /> : (
            <div className="muted small">{tr("No model has an MP plan yet. Add processes in a model's Mass Production › Plan.")}</div>
          )}
        </div>
      </section>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>{tr("Models")}<span className="cnt">{rows.length}</span></h2>
          {failed.length > 0 && <div className="wk-warn">Could not load {failed.map((r) => r.name).join(", ")}.</div>}
        </div>
        <div className="pi-sec-b wk-hm-models">
          <table className="grid">
            <thead>
              <tr>
                {["Model", "Group", "Phase", "Customer", "Build", "PM owner", "SOP target"].map((h) => <th key={h}>{tr(h)}</th>)}
                {["Plan events", "Open dev issues", "Open MP issues", "Late issues"].map((h) => <th key={h} className="num">{tr(h)}</th>)}
                <th>{tr("Updated")}</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.name} className={r.error ? "bad" : undefined}>
                  <td><button className="wk-hm-link" onClick={() => props.onOpen(r.name, "info")}>{r.name}</button></td>
                  <td className="muted">{r.path.join(" › ")}</td>
                  <td>{r.phase && <span className="wk-phase">{r.phase}</span>}</td>
                  <td>{r.customer}</td>
                  <td>{r.build}</td>
                  <td>{r.owner}</td>
                  <td>{r.sopTarget && slashDate(r.sopTarget)}</td>
                  <td className="num">{r.events || ""}</td>
                  <td className="num">{r.devOpen || ""}</td>
                  <td className="num">{r.mpOpen || ""}</td>
                  <td className="num">{r.lateIssues > 0 && <span className="wk-open">{r.lateIssues}</span>}</td>
                  <td className="muted">{r.error ? <span className="wk-warn">{r.error}</span> : r.updatedAt ? slashDate(r.updatedAt.slice(0, 10)) : ""}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

/** Per stage: the heading, the model tab a model opens on, and the text when there is nothing to show. */
const STAGE_LIST: Record<IssueStage, { title: string; tab: "devissue" | "issue"; tabName: string; none: string; noneOpen: string }> = {
  dev: { title: "Dev Issues", tab: "devissue", tabName: "Dev Issue", none: "No model has a Dev issue yet.", noneOpen: "No open Dev issue." },
  mp: { title: "MP Issues", tab: "issue", tabName: "Issues", none: "No model has an MP issue yet.", noneOpen: "No open MP issue." },
};

/**
 * The Development or Mass Production issues of every model in one list, the open ones first (or all of them); under each
 * issue its notes not done and its open sub-issues (all of them with All), each with its notes not done.
 */
function StageIssues(props: { stage: IssueStage; loads: ModelLoad[]; today: string; onOpen: (model: string, tab: "devissue" | "issue") => void }) {
  const tr = useTr();
  const [all, setAll] = useState(false);
  const list = STAGE_LIST[props.stage];
  const rows = issueRows(props.loads, props.today, props.stage);
  const issues = rows.flatMap((r) => [r, ...r.subs]);
  const open = issues.filter((r) => r.open);
  const shown = all ? rows : rows.filter(hasOpen);
  const late = open.filter((r) => r.late).length;
  const notes = (r: IssueRow, sub: boolean) => r.openNotes.map(({ no, log, late: isLateNow }) => (
    <tr key={log.id} className={"note" + (sub ? " deep" : "")}>
      <td></td>
      <td className="num">{no}</td>
      <td colSpan={2}></td>
      <td className="l"><span className="wk-hm-sub">↳ {tr(noteParts(log).line)}</span></td>
      <td className="l detail">{tr(noteParts(log).detail)}</td>
      <td>{log.owner || <span className="muted" title="Written by">{userName(log.createdBy)}</span>}</td>
      <td>{shortDate(log.date)}</td>
      <td className={isLateNow ? "late" : ""}>{log.due && shortDate(log.due)}{isLateNow && <span className="wk-late">{tr("Late")}</span>}</td>
      <td><span className="wk-st open">{tr("Open")}</span></td>
      <td colSpan={2}></td>
    </tr>
  ));
  const issueRow = (r: IssueRow, sub: boolean) => {
    const { model, issue: e, open: isOpen, late: isLateNow } = r;
    return (
      <Fragment key={model + e.id}>
        <tr className={[sub && "sub", !isOpen && "done"].filter(Boolean).join(" ") || undefined}>
          <td>{!sub && <button className="wk-hm-link" onClick={() => props.onOpen(model, list.tab)} title={`Open the model's ${list.tabName}`}>{model}</button>}</td>
          <td className="num">{r.no}</td>
          <td>{tr(e.category)}</td>
          <td>{tr(e.process)}</td>
          <td className="l">
            {sub ? <span className="wk-hm-sub">↳ <b>{tr(e.title)}</b></span> : <b>{tr(e.title)}</b>}
          </td>
          <td className="l detail">{tr(e.body)}</td>
          <td>{e.owner}</td>
          <td>{shortDate(e.date)}</td>
          <td className={isLateNow ? "late" : ""}>{e.due && shortDate(e.due)}{isLateNow && <span className="wk-late">{tr("Late")}</span>}</td>
          <td>{e.doneDate ? shortDate(e.doneDate) : ""}</td>
          <td className="l">{tr(e.finalResult)}</td>
          <td className="num">{e.logs?.length || ""}</td>
        </tr>
        {notes(r, sub)}
        {r.subs.filter((s) => all || s.open).map((s) => issueRow(s, true))}
      </Fragment>
    );
  };
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr(list.title)}<span className="cnt">{open.length}</span></h2>
        <div className="pi-legend">
          {late > 0 && <span className="wk-open">{late} {tr("late")}</span>}
          <button className={"wk-chip" + (all ? "" : " on")} onClick={() => setAll(false)}>{tr("Open")} {open.length}</button>
          <button className={"wk-chip" + (all ? " on" : "")} onClick={() => setAll(true)}>{tr("All")} {issues.length}</button>
        </div>
      </div>
      <div className="pi-sec-b wk-hm-models">
        {!shown.length ? <div className="muted small">{tr(all ? list.none : list.noneOpen)}</div> : (
          <table className="grid wk-hm-issues">
            <thead>
              <tr>
                {["Model", "No.", "Issue type", "Process", "Issue", "Issue detail", "Owner", "Registered", "Due", "Done", "Final result", "Notes"].map((h) => <th key={h}>{tr(h)}</th>)}
              </tr>
            </thead>
            <tbody>{shown.map((r) => issueRow(r, false))}</tbody>
          </table>
        )}
      </div>
    </section>
  );
}

/** A plan colour (#rrggbb) made see-through, for the light fill of task bars. */
const tint = (color: string, alpha: string) => (/^#[0-9a-f]{6}$/i.test(color) ? color + alpha : color);
const taskSpanText = (t: TaskBrief) => (t.span ? `${slashDate(t.span[0]).slice(5)}${t.span[1] !== t.span[0] ? ` – ${slashDate(t.span[1]).slice(5)}` : ""}` : "");

/** The actual days as written beside a bar: "Act. 11/18 – 11/20", "Act. 11/18 –" while still going. */
const actualText = (a: Actual) =>
  `Act. ${slashDate(a.span[0]).slice(5)}${a.going ? " –" : a.span[1] !== a.span[0] ? ` – ${slashDate(a.span[1]).slice(5)}` : ""}`;

/**
 * A bar on the timeline with its days written beside it (labelSide), and under it a thin line of the days it really took
 * (green, red when it ended or still goes after the planned end); the words go beside both.
 */
function Bar(props: {
  className: string; bar: { left: number; width: number } | null; text: string; title: string; style: CSSProperties; children?: ReactNode;
  actual?: { bar: { left: number; width: number }; of: Actual };
}) {
  const { bar, title, actual } = props;
  const both = [bar, actual?.bar].filter((b): b is { left: number; width: number } => !!b);
  const left = Math.min(...both.map((b) => b.left)), right = Math.max(...both.map((b) => b.left + b.width));
  const side = labelSide({ left, width: right - left });
  const act = actual && <span className={"wk-hm-act" + (actual.of.over ? " over" : "")}>{actualText(actual.of)}</span>;
  const words = <>{props.text}{props.text && act && " · "}{act}</>;
  return (
    <>
      {bar && (
        <span className={props.className} style={{ ...props.style, left: `${bar.left}%`, width: `${bar.width}%` }} title={title}>
          {props.children}{side === "in" && <span className="wk-hm-btext">{words}</span>}
        </span>
      )}
      {actual && (
        <span className={"wk-hm-abar" + (actual.of.over ? " over" : "") + (actual.of.going ? " going" : "")}
          style={{ left: `${actual.bar.left}%`, width: `${actual.bar.width}%` }} title={`${title} · ${actualText(actual.of)}`} />
      )}
      {side === "right" && <span className="wk-hm-blabel" style={{ left: `calc(${right}% + 6px)` }} title={title}>{words}</span>}
      {side === "left" && <span className="wk-hm-blabel" style={{ right: `calc(${100 - left}% + 6px)` }} title={title}>{words}</span>}
    </>
  );
}

/** Where the actual days of an event or task go on the timeline. */
const actualBar = (a: Actual | undefined, range: { from: string; days: number } | null) => (a && range ? { bar: barAt(a.span, range), of: a } : undefined);

/**
 * The plans on one timeline: a header line per model (what is under way, what is next, late tasks, tasks done), a bar per
 * event with how much of it is done, and a thin bar per task under it; months shaded in turn and today marked.
 */
function PlanTimeline(props: {
  kind: "dev" | "mp"; plans: PlanBrief[]; today: string; onOpen: (model: string, tab: "plan" | "setup") => void; folded?: string[];
}) {
  const mp = props.kind === "mp";
  const tr = useTr();
  // an event's tasks folded away under it (by its arrow), remembered in the browser
  const [folded, setFolded] = useState<Set<string>>(() => new Set(props.folded ?? savedFolds()));
  const toggle = (k: string) => setFolded((cur) => {
    const n = new Set(cur);
    if (!n.delete(k)) n.add(k);
    saveFolds(n);
    return n;
  });
  const range = timelineRange(props.plans, props.today);
  const months = range ? timelineMonths(range) : [];
  const todayAt = range ? barAt([props.today, props.today], range) : null;
  const todayLeft = todayAt ? `${todayAt.left + todayAt.width / 2}%` : "";
  const track = (children?: ReactNode) => (
    <div className="wk-hm-track">
      {months.map(([m, left, width], i) => <span key={m} className={"wk-hm-band" + (i % 2 ? " alt" : "")} style={{ left: `${left}%`, width: `${width}%` }} />)}
      {todayAt && <span className="wk-hm-today" style={{ left: todayLeft }} />}
      {children}
    </div>
  );

  return (
    <div className="wk-hm-plan">
      <div className="wk-hm-row head">
        <div className="wk-hm-label">{tr(mp ? "Model · line · process" : "Model · event · task")}</div>
        <div className="wk-hm-track">
          {months.map(([m, left, width], i) => (
            <span key={m} className={"wk-hm-month" + (i % 2 ? " alt" : "")} style={{ left: `${left}%`, width: `${width}%` }}>
              {width >= 4 && m.replace("-", "/")}
            </span>
          ))}
          {todayAt && <span className="wk-hm-today-l" style={{ left: todayLeft }}>{tr("Today")} {slashDate(props.today).slice(5)}</span>}
        </div>
      </div>
      {props.plans.map((p) => {
        const done = p.events.reduce((n, e) => n + e.done, 0), total = p.events.reduce((n, e) => n + e.total, 0);
        return (
          <div key={p.name} className="wk-hm-model">
            <div className="wk-hm-mhead">
              <button className="wk-hm-link" onClick={() => props.onOpen(p.name, mp ? "setup" : "plan")} title={mp ? "Open the Plan" : "Open the Dev Plan"}>{p.name}</button>
              {p.phase && <span className="wk-phase">{p.phase}</span>}
              {p.path.length > 0 && <span className="wk-hm-path">{p.path.join(" › ")}</span>}
              <span className="wk-hm-facts">
                <span className="wk-hm-fact"><i>{tr("Now")}</i>{p.now.length ? p.now.map((e) => tr(e.name)).join(", ") : "—"}</span>
                <span className="wk-hm-fact"><i>{tr("Next")}</i>{p.next ? `${tr(p.next.name)} · ${slashDate(p.next.span![0])}` : "—"}</span>
                {total > 0 && (
                  <span className="wk-hm-fact"><i>{tr(mp ? "Setup done" : "Tasks")}</i>{done}/{total}
                    <span className="wk-hm-meter"><span style={{ width: `${(done / total) * 100}%` }} /></span>
                  </span>
                )}
                {p.lateTasks > 0 && <span className="wk-open">{p.lateTasks} {tr(mp ? (p.lateTasks > 1 ? "late processes" : "late process") : p.lateTasks > 1 ? "late tasks" : "late task")}</span>}
              </span>
            </div>
            {p.events.map((e) => {
              const bar = e.span && range ? barAt(e.span, range) : null;
              const title = `${e.name}: ${spanText(e)}${e.total ? ` · ${e.done}/${e.total} tasks done` : ""}`;
              const fk = `${props.kind}|${p.name}|${e.id}`, shut = e.tasks.length > 0 && folded.has(fk);
              return (
                <div key={e.id} className="wk-hm-event">
                  <div className="wk-hm-row ev">
                    <div className="wk-hm-label">
                      {e.tasks.length > 0 ? (
                        <button className={"wk-hm-fold" + (shut ? " shut" : "")} aria-expanded={!shut} onClick={() => toggle(fk)}
                          aria-label={`${shut ? "Unfold" : "Fold"} the ${mp ? "processes" : "tasks"} of ${e.name}`}
                          title={tr(shut ? (mp ? "Show the processes" : "Show the tasks") : (mp ? "Hide the processes" : "Hide the tasks"))}>▾</button>
                      ) : <span className="wk-hm-fold none" />}
                      <span className="wk-hm-dot" style={{ background: e.color }} />
                      <span className="wk-hm-ev" title={e.name}>{tr(e.name)}</span>
                      {e.total > 0 && <span className="wk-hm-count">{e.done}/{e.total}</span>}
                      <span className={"wk-hm-st " + e.state}>{tr(STATE_LABEL[e.state])}</span>
                    </div>
                    {track((bar || e.actual) && (
                      <Bar className={"wk-hm-bar " + e.state} bar={bar} title={title} style={{ background: tint(e.color, "33"), borderColor: e.color }}
                        text={bar ? spanText(e) : ""} actual={actualBar(e.actual, range)}>
                        {e.total > 0 && <span className="wk-hm-fill" style={{ width: `${(e.done / e.total) * 100}%`, background: e.color }} />}
                      </Bar>
                    ))}
                  </div>
                  {!shut && e.tasks.filter((t) => !t.parent || !folded.has(`${fk}|${t.parent}`)).map((t) => {
                    const tb = t.span && range ? barAt(t.span, range) : null;
                    const state = t.done ? "done" : t.late ? "late" : "open";
                    // a task's sub-tasks fold away under it the same way, by "kind|model|event|task"
                    const subs = t.parent ? 0 : e.tasks.filter((c) => c.parent === t.id).length;
                    const tk = `${fk}|${t.id}`, tshut = subs > 0 && folded.has(tk);
                    return (
                      <div key={t.id} className={"wk-hm-row task " + state + (t.parent ? " sub" : "")}>
                        <div className="wk-hm-label">
                          {subs > 0 && (
                            <button className={"wk-hm-fold t" + (tshut ? " shut" : "")} aria-expanded={!tshut} onClick={() => toggle(tk)}
                              aria-label={`${tshut ? "Unfold" : "Fold"} the sub-tasks of ${t.name}`}
                              title={tr(tshut ? "Show the sub-tasks" : "Hide the sub-tasks")}>▾</button>
                          )}
                          <span className={"wk-hm-check " + state}>{t.done ? "✓" : ""}</span>
                          <span className="wk-hm-task" title={t.name}>{tr(t.name)}</span>
                          {tshut && <span className="wk-hm-count" title={tr("Sub-tasks hidden")}>+{subs}</span>}
                          {t.owner && <span className="wk-hm-owner" title="Owner">{t.owner}</span>}
                          {t.late && <span className="wk-hm-st late">{tr("Late")}</span>}
                          {!t.span && <span className="wk-hm-st undated">{tr("No date")}</span>}
                        </div>
                        {track((tb || t.actual) && (
                          <Bar className={"wk-hm-tbar " + state + (t.parent ? " sub" : "")} bar={tb} text={taskSpanText(t)} actual={actualBar(t.actual, range)} title={`${t.name}: ${taskSpanText(t)}${t.owner ? ` · ${t.owner}` : ""}${t.done ? " · done" : ""}`}
                            style={t.done ? { background: e.color, borderColor: e.color }
                              // a sub-task: a thin filled line in a lighter shade, not an outlined bar like its task
                              : t.parent ? { background: tint(e.color, "80"), borderColor: "transparent" } : { background: tint(e.color, "26"), borderColor: e.color }} />
                        ))}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}
