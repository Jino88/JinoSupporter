import { useState } from "react";
import { NumInput, TextInput } from "./fields";
import {
  WEEKDAYS,
  chainPlan,
  fmt,
  isWeekend,
  setPlanCell,
  setPlanEvent,
  signed,
  sum,
  weekday,
  type MaterialRow,
  type WipRow,
} from "./logic";
import type { DayData, MaterialDay, ModelSetup, PlanData } from "./types";

const SOURCES = ["", "GN receipt", "From PV1", "From DV", "Return"];
const diffClass = (d: number | null) => (d == null ? "" : d === 0 ? "pi-diff-ok" : "pi-diff-bad");

export default function StockPlanTab(props: {
  setup: ModelSetup;
  data: DayData;
  materials: MaterialRow[];
  wip: WipRow[];
  onData: (data: DayData) => void;
  plan: PlanData;
  onPlan: (plan: PlanData) => void;
  dates: string[];
  today: string;
  actualOf: (stageId: string) => Record<string, number>;
  onShift: (days: number) => void;
}) {
  const { setup, data, onData, plan, onPlan, dates, today } = props;
  const [lead, setLead] = useState(1);

  const setMat = (row: MaterialRow, change: Partial<MaterialDay>) => {
    const cur: MaterialDay = data.materials[row.name] ?? { opening: null, in: 0, source: "", out: null, counted: null };
    onData({ ...data, materials: { ...data.materials, [row.name]: { ...cur, ...change } } });
  };

  return (
    <>
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Materials <span className="cnt">In / Out / Stock</span></h2>
          <span className="muted small">Opening = previous day's stock (counted, else closing) · Out = Σ process input × qty per unit (from Setup); overwrite when needed.</span>
        </div>
        <div className="pi-wrap">
          <table className="pi-sm">
            <thead>
              <tr><th>Material</th><th className="n">Opening</th><th className="n">In</th><th>Source</th><th className="n">Out</th>
                <th>Used by (Setup)</th><th className="n">Closing</th><th className="n">Counted</th><th className="n">Diff</th></tr>
            </thead>
            <tbody>
              {props.materials.length === 0 && <tr><td colSpan={9} className="muted">No material — add materials in Setup.</td></tr>}
              {props.materials.map((m) => (
                <tr key={m.name}>
                  <td><b>{m.name}</b></td>
                  <td className="n">
                    <span className="pi-auto">
                      <NumInput className="wide" value={m.opening} label={`${m.name} opening`}
                        onCommit={(v) => setMat(m, { opening: v ?? 0 })} />
                      <button className={m.openingTyped ? "pi-tag-o" : "pi-tag-a"} disabled={!m.openingTyped}
                        title={m.openingTyped ? "Typed. Click to go back to the previous day's stock" : "Previous day's stock (counted, else closing)"}
                        onClick={() => setMat(m, { opening: null })}>{m.openingTyped ? "SET" : "AUTO"}</button>
                    </span>
                  </td>
                  <td className="n"><NumInput className="wide" value={m.in} label={`${m.name} in`} onCommit={(v) => setMat(m, { in: v ?? 0 })} /></td>
                  <td>
                    <select className="pi-sel" value={m.source} aria-label={`${m.name} source`} onChange={(e) => setMat(m, { source: e.target.value })}>
                      {[...new Set([m.source, ...SOURCES])].map((s) => <option key={s} value={s}>{s || "—"}</option>)}
                    </select>
                  </td>
                  <td className="n">
                    <span className="pi-auto">
                      <NumInput className="wide" value={m.out} label={`${m.name} out`} onCommit={(v) => setMat(m, { out: v ?? 0 })} />
                      <button className={m.outTyped ? "pi-tag-o" : "pi-tag-a"} disabled={!m.outTyped}
                        title={m.outTyped ? "Typed. Click to go back to AUTO" : "Σ input × qty"}
                        onClick={() => setMat(m, { out: null })}>{m.outTyped ? "SET" : "AUTO"}</button>
                    </span>
                  </td>
                  <td className="pi-used">
                    {m.uses.length ? m.uses.map((u) => <div key={u.no + u.process.id}>{u.no}. {u.process.name} × {u.qtyPerUnit}</div>) : "— not linked"}
                  </td>
                  <td className="n"><b>{fmt(m.closing)}</b></td>
                  <td className="n"><NumInput className="wide" allowEmpty value={m.counted} label={`${m.name} counted`} onCommit={(v) => setMat(m, { counted: v })} /></td>
                  <td className={"n " + diffClass(m.diff)}>{m.diff == null ? "—" : signed(m.diff)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Semi-finished (WIP)</h2>
          <span className="muted small">System = from today's flow · Counted = floor count</span>
        </div>
        <div className="pi-wrap">
          <table className="pi-sm">
            <thead><tr><th>Stage</th><th className="n">System</th><th className="n">Counted</th><th className="n">Diff</th></tr></thead>
            <tbody>
              {props.wip.map((w) => (
                <tr key={w.key}>
                  <td><b>{w.name}</b><div className="muted small">{w.hint}</div></td>
                  <td className="n">{fmt(w.system)}</td>
                  <td className="n">
                    <NumInput className="wide" allowEmpty value={w.counted} label={`${w.name} counted`}
                      onCommit={(v) => onData({ ...data, wipCount: { ...data.wipCount, [w.key]: v } })} />
                  </td>
                  <td className={"n " + diffClass(w.diff)}>{w.diff == null ? "—" : signed(w.diff)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Production plan <span className="cnt">{dates[0]} – {dates.at(-1)}</span></h2>
          <div className="pi-legend">
            <button className="btn sm" onClick={() => props.onShift(-7)}>← Week</button>
            <button className="btn sm" onClick={() => props.onShift(7)}>Week →</button>
            <span>Lead time between stages</span>
            <select className="pi-sel" value={lead} onChange={(e) => setLead(Number(e.target.value))} aria-label="Lead time">
              <option value={0}>0 day</option><option value={1}>1 day</option><option value={2}>2 days</option>
            </select>
            <button className="btn sm" onClick={() => onPlan(chainPlan(setup.stages, plan, dates, today, lead))}>
              Fill next stages from the first
            </button>
          </div>
        </div>
        <div className="pi-wrap">
          <table className="pi-plan">
            <thead>
              <tr>
                <th style={{ position: "sticky", left: 0 }}>Stage</th><th style={{ position: "sticky", left: 190 }}></th><th>Total</th>
                {dates.map((d) => (
                  <th key={d} className={d === today ? "today" : isWeekend(d) ? "we" : ""}>{d.slice(5)}<br />{WEEKDAYS[weekday(d)]}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              <tr>
                <td className="stage">Material arrival</td><td className="kind">event</td><td className="tot"></td>
                {dates.map((d) => (
                  <td key={d}><TextInput value={plan.events[d] ?? ""} label={`Event ${d}`} onCommit={(v) => onPlan(setPlanEvent(plan, d, v))} /></td>
                ))}
              </tr>
              {setup.stages.map((st) => {
                const planned = plan.values[st.id] ?? {}, act = props.actualOf(st.id);
                const tp = sum(dates.map((d) => planned[d] ?? 0)), ta = sum(dates.map((d) => act[d] ?? 0));
                const planToDate = sum(dates.filter((d) => d <= today).map((d) => planned[d] ?? 0));
                return [
                  <tr className="first" key={st.id + "p"}>
                    <td className="stage" rowSpan={3}>{st.name}<small>{st.source ? "Actual = Daily NG input" : "Actual typed here"}</small></td>
                    <td className="kind">Plan</td><td className="tot">{fmt(tp)}</td>
                    {dates.map((d) => (
                      <td key={d} className={isWeekend(d) ? "we" : ""}>
                        <NumInput allowEmpty value={planned[d] ?? null} label={`${st.name} plan ${d}`}
                          onCommit={(v) => onPlan(setPlanCell(plan, "values", st.id, d, v))} />
                      </td>
                    ))}
                  </tr>,
                  <tr key={st.id + "a"}>
                    <td className="kind">Actual</td><td className="tot">{fmt(ta)}</td>
                    {dates.map((d) => (
                      <td key={d}>
                        {!st.source && d <= today
                          ? <NumInput allowEmpty value={act[d] ?? null} label={`${st.name} actual ${d}`}
                              onCommit={(v) => onPlan(setPlanCell(plan, "actuals", st.id, d, v))} />
                          : act[d] != null ? fmt(act[d]) : ""}
                      </td>
                    ))}
                  </tr>,
                  <tr className="gap" key={st.id + "g"}>
                    <td className="kind">Gap</td><td className="tot">{signed(sum(dates.filter((d) => d <= today).map((d) => act[d] ?? 0)) - planToDate)}</td>
                    {dates.map((d) => {
                      if (act[d] == null || d > today) return <td key={d}></td>;
                      const gp = act[d] - (planned[d] ?? 0);
                      return <td key={d} className={gp < 0 ? "neg" : "pos"}>{signed(gp)}</td>;
                    })}
                  </tr>,
                ];
              })}
            </tbody>
          </table>
        </div>
      </section>
    </>
  );
}
