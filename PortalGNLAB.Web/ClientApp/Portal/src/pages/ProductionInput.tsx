import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { formatPublished } from "../format";
import { downloadBlob } from "../toolsApi";
import { productionApi } from "./productionInput/api";
import DailyTab, { type DayView } from "./productionInput/DailyTab";
import {
  actuals,
  addDays,
  afterOverBefore,
  calc,
  dateRange,
  fmt,
  localToday,
  materialRows,
  ngSheet,
  ngTotals,
  normalizeProcess,
  pct,
  planSheet,
  prevOutFrom,
  stageActuals,
  summarize,
  weekStart,
  wipRows,
  type DayEntry,
} from "./productionInput/logic";
import ReportsTab from "./productionInput/ReportsTab";
import SetupTab, { BUILDS } from "./productionInput/SetupTab";
import StockPlanTab from "./productionInput/StockPlanTab";
import { copySetup, emptySetup, x316bSetup } from "./productionInput/template";
import {
  emptyDayData,
  emptyPlanData,
  type Day,
  type DayData,
  type DayStatus,
  type DaySummary,
  type Model,
  type ModelSetup,
  type ModelSummary,
  type PlanData,
  type Sheet,
  type Slot,
} from "./productionInput/types";
import "./productionInput/productionInput.css";

type Tab = "setup" | "proc" | "stock" | "rep";
const WINDOW = 14;
const TEMPLATE = "template:x316b";

const normalizeDay = (data: Partial<DayData> | null | undefined): DayData => ({ ...emptyDayData(), ...(data ?? {}) });
const normalizePlan = (data: Partial<PlanData> | null | undefined): PlanData => ({ ...emptyPlanData(), ...(data ?? {}) });
const normalizeSetup = (s: Partial<ModelSetup> | null | undefined): ModelSetup =>
  ({ processes: (s?.processes ?? []).map(normalizeProcess), materials: s?.materials ?? [], stages: s?.stages ?? [] });

export default function ProductionInput() {
  const [models, setModels] = useState<ModelSummary[]>([]);
  const [model, setModel] = useState<Model | null>(null);
  const [setup, setSetup] = useState<ModelSetup>(emptySetup);
  const [modelDirty, setModelDirty] = useState(false);
  const [date, setDate] = useState(localToday);
  const [day, setDay] = useState<Day | null>(null);
  const [previous, setPrevious] = useState<Day | null>(null);
  const [data, setData] = useState<DayData>(emptyDayData);
  const [dayDirty, setDayDirty] = useState(false);
  const [planRevision, setPlanRevision] = useState(0);
  const [plan, setPlan] = useState<PlanData>(emptyPlanData);
  const [planDirty, setPlanDirty] = useState(false);
  const [windowStart, setWindowStart] = useState(() => weekStart(localToday()));
  const [windowDays, setWindowDays] = useState<Day[]>([]);
  const [tab, setTab] = useState<Tab>("proc");
  const [slot, setSlot] = useState<Slot>("day");
  const [dayView, setDayView] = useState<DayView>("date");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [newName, setNewName] = useState("");
  const [newBuild, setNewBuild] = useState("PV2");
  const [newCopy, setNewCopy] = useState(TEMPLATE);
  const load = useRef(0);

  const dirty = modelDirty || dayDirty || planDirty;
  const dates = useMemo(() => dateRange(windowStart, WINDOW), [windowStart]);

  const fail = (e: unknown) => setError((e as Error).message);
  const say = (message: string) => { setNotice(message); setError(""); window.setTimeout(() => setNotice(""), 2500); };
  const confirmDiscard = () => !dirty || window.confirm("Discard the unsaved changes?");

  const refreshModels = useCallback(() => productionApi.models().then((list) => { setModels(list); return list; }), []);

  useEffect(() => {
    refreshModels()
      .then((list) => { if (list.length) openModel(list[0].id); else setTab("setup"); })
      .catch(fail);
  }, []);

  async function openModel(id: string) {
    const ticket = ++load.current;
    try {
      const [m, p] = await Promise.all([productionApi.model(id), productionApi.plan(id)]);
      if (ticket !== load.current) return;
      const s = normalizeSetup(m.setup);
      setModel({ ...m, setup: s });
      setSetup(s);
      setModelDirty(false);
      setPlan(normalizePlan(p.data));
      setPlanRevision(p.revision);
      setPlanDirty(false);
      if (!s.processes.length) setTab("setup");
    } catch (e) { fail(e); }
  }

  // day on screen + the saved days of the plan window
  useEffect(() => {
    if (!model) return;
    const ticket = load.current;
    Promise.all([productionApi.day(model.id, date), productionApi.days(model.id, windowStart, addDays(windowStart, WINDOW - 1))])
      .then(([d, list]) => {
        if (ticket !== load.current) return;
        setDay(d.day);
        setPrevious(d.previous);
        setData(normalizeDay(d.day?.data));
        setDayDirty(false);
        setWindowDays(list);
      })
      .catch(fail);
  }, [model?.id, date, windowStart]);

  const prevOut = useMemo(() => prevOutFrom(previous?.summary), [previous]);
  const rows = useMemo(() => calc(setup, data, prevOut), [setup, data, prevOut]);
  const mats = useMemo(() => materialRows(setup, rows, data, previous?.summary), [setup, rows, data, previous]);
  const wip = useMemo(() => wipRows(setup, rows, data), [setup, rows, data]);
  const live = useMemo(() => actuals(setup, rows, data), [setup, rows, data]);
  const summary = useMemo(() => summarize(setup, data, rows, prevOut, mats), [setup, data, rows, prevOut, mats]);
  const totals = ngTotals(setup, rows);
  const badCells = afterOverBefore(setup, data);

  const summaries = useMemo(() => {
    const map: Record<string, DaySummary> = {};
    for (const d of windowDays) if (d.summary) map[d.date] = d.summary;
    return map;
  }, [windowDays]);
  const actualOf = (stageId: string) => {
    const stage = setup.stages.find((s) => s.id === stageId);
    return stage ? stageActuals(stage, plan, summaries, { date, actuals: live }) : {};
  };
  const planFor = (source: string) => {
    const stage = setup.stages.find((s) => s.source === source);
    return stage ? plan.values[stage.id]?.[date] ?? null : null;
  };

  // entered days of the window; the day on screen counts with its unsaved input
  const entries = useMemo(() => {
    const list: DayEntry[] = windowDays
      .filter((d) => d.date !== date)
      .map((d) => ({ date: d.date, data: normalizeDay(d.data), summary: d.summary }));
    if ((day || dayDirty) && dates.includes(date)) list.push({ date, data, summary });
    return list.sort((a, b) => a.date.localeCompare(b.date));
  }, [windowDays, date, day, dayDirty, data, summary, dates]);
  const statusOf = useMemo(() => {
    const map: Record<string, DayStatus> = {};
    for (const d of windowDays) map[d.date] = d.status;
    if (day) map[date] = day.status;
    return map;
  }, [windowDays, day, date]);

  const sheets = useMemo(() => {
    const name = model ? `${model.name} ${model.build}` : "";
    return {
      plan: planSheet(name, setup.stages, plan, dates, (st) => stageActuals(st, plan, summaries, { date, actuals: live })),
      ng: ngSheet(name, setup, entries),
      ngDays: entries.length,
    };
  }, [model, entries, date, setup, plan, dates, summaries, live]);

  function editSetup(next: ModelSetup) { setSetup(next); setModelDirty(true); }
  function editData(next: DayData) { setData(next); setDayDirty(true); }
  function editPlan(next: PlanData) { setPlan(next); setPlanDirty(true); }

  async function save(status?: DayStatus) {
    if (!model) return;
    if (status === "Submitted" && badCells) { setError(`${badCells} cell(s): After is larger than Before.`); return; }
    setBusy(true);
    setError("");
    try {
      if (modelDirty) {
        const saved = await productionApi.saveModel({ ...model, setup });
        setModel({ ...saved, setup: normalizeSetup(saved.setup) });
        setModelDirty(false);
        await refreshModels();
      }
      if (planDirty) {
        const saved = await productionApi.savePlan(model.id, plan, planRevision);
        setPlanRevision(saved.revision);
        setPlanDirty(false);
      }
      const nextStatus = status ?? day?.status ?? "Draft";
      if (dayDirty || (status && status !== day?.status)) {
        const saved = await productionApi.saveDay(model.id, date, { status: nextStatus, data, summary, revision: day?.revision ?? 0 });
        setDay(saved);
        setDayDirty(false);
        setWindowDays((list) => [...list.filter((d) => d.date !== date), saved].filter((d) => dates.includes(d.date)));
      }
      say(status === "Submitted" ? "Day submitted." : "Saved.");
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  }

  async function addModel() {
    const name = newName.trim();
    if (!name) { setError("Type a model name."); return; }
    if (!confirmDiscard()) return;
    setBusy(true);
    try {
      let start: ModelSetup;
      if (newCopy === TEMPLATE) start = x316bSetup();
      else if (newCopy) start = copySetup(normalizeSetup((await productionApi.model(newCopy)).setup));
      else start = emptySetup();
      const created = await productionApi.createModel({ name, build: newBuild, setup: start });
      setNewName("");
      await refreshModels();
      await openModel(created.id);
      setTab("setup");
      say("Model added — check its processes, NG items and materials.");
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  }

  async function removeModel() {
    if (!model || !window.confirm(`Delete ${model.name} ${model.build} with all its daily input and plan?`)) return;
    setBusy(true);
    try {
      await productionApi.deleteModel(model.id);
      const list = await refreshModels();
      setModel(null);
      setModelDirty(false); setDayDirty(false); setPlanDirty(false);
      if (list.length) await openModel(list[0].id);
      say("Model deleted.");
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  }

  async function exportSheets(list: Sheet[]) {
    if (!model) return;
    setBusy(true);
    try {
      const fileName = `${model.name} ${model.build} ${dates[0]}~${dates.at(-1)}.xlsx`;
      downloadBlob(await productionApi.exportXlsx(fileName, list), fileName);
    } catch (e) {
      fail(e);
    } finally {
      setBusy(false);
    }
  }

  const changeDate = (next: string) => {
    if (!next || next === date || !confirmDiscard()) return;
    setDate(next);
    setWindowStart(weekStart(next));
    setDayView("date");
  };
  // a date picked in the Daily NG list keeps the window where it is
  const pickDate = (next: string) => {
    setDayView("date");
    if (next !== date && confirmDiscard()) setDate(next);
  };
  const shiftWindow = (n: number) => {
    if (!dayDirty || window.confirm("Discard the unsaved daily input?")) setWindowStart(addDays(windowStart, n));
  };

  const status = !day ? <span className="pi-st none">Not entered</span>
    : day.status === "Submitted" ? <span className="pi-st ok">Submitted {formatPublished(day.updatedAt)}</span>
    : <span className="pi-st">Draft · not submitted</span>;
  const kp = (key: string, v: string, n: string, cls = "") => (
    <div className="pi-kp"><div className="pi-kp-k">{key}</div><div className={"pi-kp-v " + cls}>{v}</div><div className="pi-kp-n">{n}</div></div>
  );
  const planNote = (source: string, tail: string) => `Plan ${fmt(planFor(source))} · ${tail}`;

  return (
    <div className="tools-page pi-page">
      <header className="page-head">
        <h1>Production Daily Input</h1>
        <p className="muted">Register the model first, then its processes, NG items and the materials each process uses. Daily entry is only NG counts; inputs, NG rates, stock and plan actuals follow automatically.</p>
      </header>
      {error && <div className="alert" role="alert">{error}</div>}

      <div className="pi-lay">
        <aside className="card pi-models">
          <h2>Models</h2>
          <div className="pi-hint">Step 1 — register the model. Each model keeps its own processes, NG items and materials.</div>
          {models.length === 0 && <div className="muted small">No model yet.</div>}
          {models.map((m) => (
            <button key={m.id} className={"pi-mi " + (m.id === model?.id ? "on" : "")}
              onClick={() => m.id !== model?.id && confirmDiscard() && openModel(m.id)}>
              <b>{m.name} <span>{m.build}</span></b>
              <span>{m.processCount} processes · {m.materialCount} materials</span>
            </button>
          ))}
          <div className="pi-madd">
            <label htmlFor="pi-new-model">New model</label>
            <input id="pi-new-model" className="inp sm" placeholder="e.g. MSM-X516B" value={newName}
              onChange={(e) => setNewName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addModel()} />
            <select className="inp sm" value={newBuild} onChange={(e) => setNewBuild(e.target.value)} aria-label="Build">
              {BUILDS.map((b) => <option key={b}>{b}</option>)}
            </select>
            <select className="inp sm" value={newCopy} onChange={(e) => setNewCopy(e.target.value)} aria-label="Start from">
              <option value={TEMPLATE}>Start from X316B PV2 workbook</option>
              <option value="">Start empty</option>
              {models.map((m) => <option key={m.id} value={m.id}>Copy setup of {m.name} {m.build}</option>)}
            </select>
            <button className="btn sm primary" disabled={busy} onClick={addModel}>+ Add model</button>
          </div>
        </aside>

        {!model ? (
          <section className="card pi-empty">Add a model on the left to start.</section>
        ) : (
          <div>
            <section className="card pi-hdr">
              <div className="pi-field"><label>Model</label><div className="pi-mname">{model.name} · {model.build}</div></div>
              <div className="pi-field"><label htmlFor="pi-date">Date</label>
                <input id="pi-date" className="inp" type="date" value={date} onChange={(e) => changeDate(e.target.value)} /></div>
              <div className="pi-field"><label>Entered by</label><div className="pi-mname" style={{ fontWeight: 500, fontSize: 13 }}>{day?.enteredBy || "—"}</div></div>
              <div className="pi-field"><label>Status</label><div className="pi-mname">{status}</div></div>
              <div className="pi-right">
                {notice && <span className="pi-msg">{notice}</span>}
                {dirty && <span className="pi-dirty">Unsaved changes</span>}
                <button className="btn sm" disabled={busy || !dirty} onClick={() => save()}>Save</button>
                <button className="btn sm primary" disabled={busy} onClick={() => save("Submitted")}>Submit day</button>
              </div>
            </section>

            <section className="card pi-kpis">
              {kp("Sub 12h actual", fmt(live.sub), planNote("sub", "into Sub Dry"), "good")}
              {kp("Main 12h actual", fmt(live.main), planNote("main", "into Main Dry"), "good")}
              {kp("Function OK", fmt(live.func), planNote("func", "last function test output"))}
              {kp("Visual OK", fmt(live.vis), planNote("vis", "Packing input"))}
              {kp("NG process", `${pct(totals.processBefore)} / ${pct(totals.processAfter)}`, "Before / After rework", totals.processAfter > 0.03 ? "bad" : "")}
              {kp("NG function", `${pct(totals.functionBefore)} / ${pct(totals.functionAfter)}`, "Function-group processes", totals.functionAfter > 0.04 ? "bad" : "")}
            </section>

            <nav className="pi-tabs">
              <button className={tab === "setup" ? "on" : ""} onClick={() => setTab("setup")}>1. Setup <span className="muted small">process · NG · material</span></button>
              <button className={tab === "proc" ? "on" : ""} onClick={() => setTab("proc")}>2. Daily NG input<span className="cnt">{setup.processes.length}</span></button>
              <button className={tab === "stock" ? "on" : ""} onClick={() => setTab("stock")}>3. Stock &amp; Plan</button>
              <button className={tab === "rep" ? "on" : ""} onClick={() => setTab("rep")}>4. Reports (Excel forms)</button>
            </nav>

            {tab === "setup" && (
              <SetupTab name={model.name} build={model.build} setup={setup} onSetup={editSetup}
                onModel={(change) => { setModel({ ...model, ...change }); setModelDirty(true); }}
                onDelete={removeModel} onError={setError} />
            )}
            {tab === "proc" && (
              <DailyTab setup={setup} data={data} rows={rows} slot={slot} onSlot={setSlot} onData={editData}
                view={dayView} date={date} dates={dates} entries={entries} statusOf={statusOf} unsaved={dayDirty}
                onPick={pickDate} onTotal={() => setDayView("total")} onShift={shiftWindow} />
            )}
            {tab === "stock" && (
              <StockPlanTab setup={setup} data={data} materials={mats} wip={wip} onData={editData}
                plan={plan} onPlan={editPlan} dates={dates} today={date} actualOf={actualOf} onShift={shiftWindow} />
            )}
            {tab === "rep" && <ReportsTab planSheet={sheets.plan} ngSheet={sheets.ng} ngDays={sheets.ngDays} busy={busy} onExport={exportSheets} />}
          </div>
        )}
      </div>
    </div>
  );
}
