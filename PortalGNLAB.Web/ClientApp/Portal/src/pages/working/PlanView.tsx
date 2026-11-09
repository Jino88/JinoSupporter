import { useEffect, useMemo, useState } from "react";
import { ApiError } from "../../api";
import { productionApi } from "../productionInput/api";
import { WEEKDAYS, addDays, dateRange, fmt, isWeekend, localToday, normalizeProcess, signed, stageActuals, sum, weekStart, weekday } from "../productionInput/logic";
import { emptyPlanData, type Day, type DaySummary, type ModelSetup, type ModelSummary, type PlanData } from "../productionInput/types";
import "../productionInput/productionInput.css";

const WINDOW = 14;

/** Production Input models whose name is this model's name (case-insensitive), one per build. */
export const matchingModels = (models: ModelSummary[], name: string) =>
  models.filter((m) => m.name.trim().toLowerCase() === name.trim().toLowerCase());

/**
 * Read-only copy of the Production Input "Production plan" table (StockPlanTab) for one model:
 * Plan / Actual / Gap per stage over two weeks. Editing stays in Production Input.
 */
export default function PlanView({ modelName }: { modelName: string }) {
  const [models, setModels] = useState<ModelSummary[] | null>(null);
  const [buildId, setBuildId] = useState("");
  const [setup, setSetup] = useState<ModelSetup | null>(null);
  const [plan, setPlan] = useState<PlanData>(emptyPlanData);
  const [days, setDays] = useState<Day[]>([]);
  const [start, setStart] = useState(() => weekStart(localToday()));
  const [error, setError] = useState("");
  const today = localToday();
  const dates = useMemo(() => dateRange(start, WINDOW), [start]);

  useEffect(() => {
    productionApi.models().then(setModels).catch((e: Error) => {
      setModels([]);
      setError(e instanceof ApiError && e.status === 403 ? "You do not have the Production Input menu, so the plan cannot be shown." : e.message);
    });
  }, []);

  const matches = useMemo(() => matchingModels(models ?? [], modelName), [models, modelName]);
  useEffect(() => { setBuildId(matches[0]?.id ?? ""); }, [matches]);

  useEffect(() => {
    if (!buildId) { setSetup(null); return; }
    let alive = true;
    Promise.all([productionApi.model(buildId), productionApi.plan(buildId)])
      .then(([m, p]) => {
        if (!alive) return;
        setSetup({ processes: (m.setup?.processes ?? []).map(normalizeProcess), materials: m.setup?.materials ?? [], stages: m.setup?.stages ?? [] });
        setPlan({ ...emptyPlanData(), ...(p.data ?? {}) });
      })
      .catch((e: Error) => alive && setError(e.message));
    return () => { alive = false; };
  }, [buildId]);

  useEffect(() => {
    if (!buildId) return;
    let alive = true;
    productionApi.days(buildId, dates[0], dates[dates.length - 1])
      .then((list) => alive && setDays(list))
      .catch((e: Error) => alive && setError(e.message));
    return () => { alive = false; };
  }, [buildId, dates]);

  const summaries = useMemo(() => {
    const map: Record<string, DaySummary> = {};
    for (const d of days) if (d.summary) map[d.date] = d.summary;
    return map;
  }, [days]);

  if (error) return <div className="alert" role="alert">{error}</div>;
  if (models === null) return <div className="muted">Loading Production Input…</div>;
  if (!matches.length)
    return (
      <div className="wk-empty">
        <b>No Production Input model named {modelName}.</b>
        <div className="muted">Register it in Production Input with the same name; its plan then shows here.</div>
        <a className="btn sm" href="/app/production/input">Open Production Input</a>
      </div>
    );

  const open = `/app/production/input?model=${encodeURIComponent(buildId)}&tab=stock`;
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Production plan <span className="cnt">{dates[0]} – {dates[dates.length - 1]}</span></h2>
        <div className="pi-legend">
          {matches.length > 1 && (
            <select className="pi-sel" value={buildId} onChange={(e) => setBuildId(e.target.value)} aria-label="Build">
              {matches.map((m) => <option key={m.id} value={m.id}>{m.build || "(no build)"}</option>)}
            </select>
          )}
          {matches.length === 1 && <span className="muted">Build {matches[0].build || "—"}</span>}
          <button className="btn sm" onClick={() => setStart(addDays(start, -7))}>← Week</button>
          <button className="btn sm" onClick={() => setStart(addDays(start, 7))}>Week →</button>
          <a className="btn sm" href={open}>Edit in Production Input</a>
        </div>
      </div>
      {!setup ? <div className="pi-sec-b muted">Loading plan…</div> : !setup.stages.length ? (
        <div className="pi-sec-b muted">This model has no plan stages yet. Add them in Production Input → Setup.</div>
      ) : (
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
                {dates.map((d) => <td key={d} className="wk-ev">{plan.events[d] ?? ""}</td>)}
              </tr>
              {setup.stages.map((st) => {
                const planned = plan.values[st.id] ?? {};
                const act = stageActuals(st, plan, summaries, null);
                const past = dates.filter((d) => d <= today);
                const tp = sum(dates.map((d) => planned[d] ?? 0)), ta = sum(dates.map((d) => act[d] ?? 0));
                const gapToDate = sum(past.map((d) => act[d] ?? 0)) - sum(past.map((d) => planned[d] ?? 0));
                return [
                  <tr className="first" key={st.id + "p"}>
                    <td className="stage" rowSpan={3}>{st.name}<small>{st.source ? "Actual = Daily NG input" : "Actual typed in plan"}</small></td>
                    <td className="kind">Plan</td><td className="tot">{fmt(tp)}</td>
                    {dates.map((d) => <td key={d} className={isWeekend(d) ? "we" : ""}>{planned[d] != null ? fmt(planned[d]) : ""}</td>)}
                  </tr>,
                  <tr key={st.id + "a"}>
                    <td className="kind">Actual</td><td className="tot">{fmt(ta)}</td>
                    {dates.map((d) => <td key={d}>{act[d] != null ? fmt(act[d]) : ""}</td>)}
                  </tr>,
                  <tr className="gap" key={st.id + "g"}>
                    <td className="kind">Gap</td><td className="tot">{signed(gapToDate)}</td>
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
      )}
    </section>
  );
}
