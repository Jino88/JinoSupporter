import { useEffect, useState, type ReactNode } from "react";
import { workingApi } from "./api";
import {
  barAt, devPlans, modelRow, slashDate, timelineMonths, timelineRange, type EventBrief, type ModelLoad, type PlanBrief,
} from "./homeSummary";

const STATE_LABEL: Record<EventBrief["state"], string> = { done: "Done", now: "Now", upcoming: "Next", late: "Late", undated: "No date" };
/** Models loaded at the same time, so a long tree does not send every request at once. */
const AT_ONCE = 6;

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
        const [model, entries] = await Promise.all([workingApi.model(name), workingApi.entries(name)]);
        out[i] = { name, path, model, entries, error: "" };
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
  onOpen: (model: string, tab: "info" | "plan") => void;
  /** Already loaded (tests and the page reopening); loaded here when left out. */
  initial?: ModelLoad[];
}) {
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
  const { plans, noPlan } = devPlans(loads, props.today);
  const rows = loads.map((l) => modelRow(l, props.today));
  const failed = rows.filter((r) => r.error);

  return (
    <div className="wk-home">
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Development plans<span className="cnt">{plans.length}</span></h2>
          <div className="pi-legend">
            <span>Models in development (not MP) with a Dev Plan · today {slashDate(props.today)}</span>
          </div>
        </div>
        <div className="pi-sec-b">
          {plans.length ? <PlanTimeline plans={plans} today={props.today} onOpen={props.onOpen} /> : (
            <div className="muted small">No model in development has a Dev Plan yet. Add events in a model's Development › Dev Plan.</div>
          )}
          {noPlan.length > 0 && (
            <div className="wk-hm-noplan">
              <span className="muted small">In development without a plan:</span>
              {noPlan.map((m) => <button key={m} className="wk-chip" onClick={() => props.onOpen(m, "plan")}>{m}</button>)}
            </div>
          )}
        </div>
      </section>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Models<span className="cnt">{rows.length}</span></h2>
          {failed.length > 0 && <div className="wk-warn">Could not load {failed.map((r) => r.name).join(", ")}.</div>}
        </div>
        <div className="pi-sec-b wk-hm-models">
          <table className="grid">
            <thead>
              <tr>
                <th>Model</th><th>Group</th><th>Phase</th><th>Customer</th><th>Build</th><th>PM owner</th><th>SOP target</th>
                <th className="num">Plan events</th><th className="num">Open dev issues</th><th className="num">Open MP issues</th><th className="num">Late issues</th><th>Updated</th>
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

/** The plans on one timeline: a line per model with what is under way and next, a bar per event under it, today marked. */
function PlanTimeline(props: { plans: PlanBrief[]; today: string; onOpen: (model: string, tab: "plan") => void }) {
  const range = timelineRange(props.plans, props.today);
  const months = range ? timelineMonths(range) : [];
  const todayAt = range ? barAt([props.today, props.today], range) : null;
  const track = (children?: ReactNode) => (
    <div className="wk-hm-track">
      {months.map(([m, left, width]) => <span key={m} className="wk-hm-grid" style={{ left: `${left}%`, width: `${width}%` }} />)}
      {todayAt && <span className="wk-hm-today" style={{ left: `${todayAt.left + todayAt.width / 2}%` }} />}
      {children}
    </div>
  );

  return (
    <div className="wk-hm-plan">
      <div className="wk-hm-row head">
        <div className="wk-hm-label" />
        <div className="wk-hm-track">
          {months.map(([m, left, width]) => (
            <span key={m} className="wk-hm-month" style={{ left: `${left}%`, width: `${width}%` }}>{m.replace("-", "/")}</span>
          ))}
          {todayAt && <span className="wk-hm-today-l" style={{ left: `${todayAt.left + todayAt.width / 2}%` }}>Today</span>}
        </div>
      </div>
      {props.plans.map((p) => (
        <div key={p.name} className="wk-hm-model">
          <div className="wk-hm-row">
            <div className="wk-hm-label">
              <button className="wk-hm-link" onClick={() => props.onOpen(p.name, "plan")} title="Open the Dev Plan">{p.name}</button>
              {p.phase && <span className="wk-phase">{p.phase}</span>}
            </div>
            <div className="wk-hm-sum">
              <span>Now: <b>{p.now.length ? p.now.map((e) => e.name).join(", ") : "—"}</b></span>
              <span>Next: <b>{p.next ? `${p.next.name} ${slashDate(p.next.span![0])}` : "—"}</b></span>
              {p.lateTasks > 0 && <span className="wk-open">{p.lateTasks} late task{p.lateTasks > 1 ? "s" : ""}</span>}
            </div>
          </div>
          {p.events.map((e) => {
            const bar = e.span && range ? barAt(e.span, range) : null;
            return (
              <div key={e.id} className="wk-hm-row ev">
                <div className="wk-hm-label">
                  <span className="wk-hm-dot" style={{ background: e.color }} />
                  <span className="wk-hm-ev" title={e.name}>{e.name}</span>
                  <span className={"wk-hm-st " + e.state}>{STATE_LABEL[e.state]}</span>
                </div>
                {track(bar && (
                  <span className={"wk-hm-bar " + e.state} style={{ left: `${bar.left}%`, width: `${bar.width}%`, background: e.color }}
                    title={`${e.name}: ${spanText(e)}${e.total ? ` · ${e.done}/${e.total} tasks done` : ""}`}>
                    {spanText(e)}{e.total > 0 && ` · tasks ${e.done}/${e.total}`}
                  </span>
                ))}
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}
