import { Fragment, useEffect, useRef, useState } from "react";
import DateBox from "./DateBox";
import { Gantt, dates } from "./DevPlan";
import { useTr } from "./i18n";
import { workingApi } from "./api";
import IssueTable from "./IssueTable";
import { blankEntry, isOpenIssue, issuesBySeq } from "./logic";
import {
  issueProcessOf, lineOf, newSetupRow, parseSetupPaste, rate, rowsInOrder, SETUP_STATES, setupApi, setupBatch, setupEvents, setupFieldsOf, setupProblem, setupSummary,
  shortDate, tableToTsv,
} from "./setup";
import { calendarDays, eventSpan, planDays, progress, taskSpan } from "./plan";
import type { PastedSetup, SetupFields, SetupRow, SetupState } from "./setup";
import type { Entry } from "./types";
import "./setup.css";

/** Columns of the process table (toggle … actions), for the rows that span it. */
const COLS = 13;

const stateClass = (s: SetupState) => (s === "O" ? "ok" : s === "X" ? "no" : s === "Δ" ? "now" : "");

function StateSelect(props: { value: SetupState; label: string; onChange: (v: SetupState) => void }) {
  return (
    <select className="inp sm su-state-in" aria-label={props.label} value={props.value} onChange={(e) => props.onChange(e.target.value as SetupState)}>
      <option value="">–</option>
      {/* the sign alone keeps the column narrow; the legend above says what it means */}
      {SETUP_STATES.map((s) => <option key={s.value} value={s.value} title={s.label}>{s.value}</option>)}
    </select>
  );
}

/** The input cells of a row being added or edited, in the process table's columns (No. … Status). */
function SetupCells(props: { value: SetupFields; onChange: (f: SetupFields) => void; autoFocus?: boolean }) {
  const f = props.value;
  const set = <K extends keyof SetupFields>(k: K, v: SetupFields[K]) => props.onChange({ ...f, [k]: v });
  const date = (k: "machinePlan" | "setupPlan" | "setupFinish", label: string) => (
    <td><DateBox className="sm su-date" label={label} value={f[k]} onChange={(v) => set(k, v)} /></td>
  );
  return (
    <>
      <td><input className="inp sm su-no" type="number" min={1} autoFocus={props.autoFocus} aria-label="No." value={f.no || ""}
        onChange={(e) => set("no", e.target.value === "" ? 0 : Number(e.target.value))} /></td>
      <td><input className="inp sm su-no" aria-label="No. layout" value={f.layoutNo} onChange={(e) => set("layoutNo", e.target.value)} /></td>
      <td><input className="inp sm su-line" aria-label="Line" list="su-lines" value={f.line} onChange={(e) => set("line", e.target.value)} /></td>
      <td className="l"><input className="inp sm su-proc" aria-label="Process name" value={f.processName} onChange={(e) => set("processName", e.target.value)} /></td>
      {date("machinePlan", "Machine plan")}
      <td><StateSelect label="Machine state" value={f.machineState} onChange={(v) => set("machineState", v)} /></td>
      <td><input className="inp sm su-note" aria-label="Machine note" value={f.machineNote} onChange={(e) => set("machineNote", e.target.value)} /></td>
      {date("setupPlan", "Setup plan")}
      {date("setupFinish", "Setup finish")}
      <td><StateSelect label="Setup status" value={f.setupStatus} onChange={(v) => set("setupStatus", v)} /></td>
    </>
  );
}

/** A saved row being edited, in place. */
function SetupEditRow(props: { value: SetupFields; busy: boolean; onSave: (f: SetupFields) => void; onCancel: () => void }) {
  const [f, setF] = useState(props.value);
  const [shown, setShown] = useState(false);
  const problem = setupProblem(f);
  const save = () => { setShown(true); if (!problem) props.onSave({ ...f, processName: f.processName.trim() }); };
  return (
    <tr className="su-edit" onKeyDown={(e) => { if (e.key === "Escape") props.onCancel(); if (e.key === "Enter" && e.target instanceof HTMLInputElement) save(); }}>
      <td></td>
      <SetupCells value={f} onChange={setF} autoFocus />
      <td className="l su-msg">
        {shown && problem ? <span className="su-err">{problem}</span> : <span className="muted small">Problems are edited below the row.</span>}
      </td>
      <td className="act">
        <button className="btn sm primary" disabled={props.busy} onClick={save}>Save</button>
        <button className="btn sm" disabled={props.busy} onClick={props.onCancel}>Cancel</button>
      </td>
    </tr>
  );
}

/** Rows copied from the Excel sheet, checked before they are added. */
function PasteBox(props: { today: string; busy: boolean; onImport: (p: PastedSetup, withProblems: boolean) => void; onCancel: () => void }) {
  const [text, setText] = useState("");
  const [withProblems, setWithProblems] = useState(true);
  const parsed = text.trim() ? parseSetupPaste(text, props.today) : null;
  const notes = parsed?.rows.filter((r) => r.note).length ?? 0;
  return (
    <div className="su-paste">
      <div className="muted small">
        In Excel, select the rows from <b>No</b> to <b>Note</b> (11 columns: No, No. Layout, LINE, Process Name, Plan, State, Note, Plan, Finish,
        Status, Note), copy them and paste here. Header lines are skipped and an empty LINE takes the one above.
      </div>
      <textarea className="inp wk-ta su-paste-ta" aria-label="Rows from Excel" placeholder="Paste the rows here (Ctrl+V)" value={text}
        onChange={(e) => setText(e.target.value)}
        onPaste={(e) => {
          // the plain text of some copies has no tabs; the HTML table Excel puts next to it keeps every column
          const plain = e.clipboardData.getData("text/plain");
          const tsv = plain.includes("\t") ? "" : tableToTsv(e.clipboardData.getData("text/html"));
          if (!tsv) return;
          e.preventDefault();
          setText((t) => (t.trim() ? t.replace(/\s*$/, "\n") : "") + tsv);
        }} />
      <div className="su-paste-f">
        {parsed?.problem ? <span className="su-err">{parsed.problem}</span>
          : parsed ? <span>{parsed.rows.length} rows{notes ? `, ${notes} with a note` : ""}.</span> : <span className="muted">Nothing pasted yet.</span>}
        <label className="wk-done">
          <input type="checkbox" checked={withProblems} onChange={(e) => setWithProblems(e.target.checked)} /> Add each note as an open problem
        </label>
        <button className="btn sm primary" disabled={props.busy || !parsed || !!parsed.problem} onClick={() => parsed && props.onImport(parsed, withProblems)}>
          Add {parsed?.rows.length ?? 0} rows
        </button>
        <button className="btn sm" disabled={props.busy} onClick={props.onCancel}>Cancel</button>
      </div>
    </div>
  );
}

/**
 * Plan of Mass Production, shown like Dev Plan: a line per row of the calendar (its processes from their machine plan to
 * their setup plan), and under it the line picked: its processes on a calendar and in a table with the machine line plan /
 * state and the setup plan / finish / status. Each process opens into its Issue (Problem) list: issues of this model raised against the row, with the same columns,
 * daily notes and pictures as in Issues (they are listed there too).
 */
export default function SetupPlan(props: {
  modelName: string;
  modelPhase: string;
  entries: Entry[];
  today: string;
  onChanged: (list: Entry[]) => void;
  onError: (e: unknown) => void;
  /** Rows to show instead of loading them (tests). */
  initial?: SetupRow[];
}) {
  const tr = useTr();
  const { modelName, entries, today } = props;
  const [rows, setRows] = useState<SetupRow[] | null>(props.initial ?? null);
  const [editing, setEditing] = useState<string | null>(null); // id of the saved row being edited
  const [adding, setAdding] = useState<{ key: number; f: SetupFields }[]>([]); // new rows, saved together
  const [addTried, setAddTried] = useState(false);
  const addKey = useRef(0);
  const [pasting, setPasting] = useState(false);
  const [open, setOpen] = useState<Set<string>>(new Set());
  const [line, setLine] = useState(""); // the line picked (the first one until another is)
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (props.initial) return;
    let live = true;
    setupApi.rows(modelName).then((r) => { if (live) setRows(r); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [modelName]);

  if (rows === null) return <div className="muted">Loading the plan…</div>;

  const ordered = rowsInOrder(rows);
  const summary = setupSummary(rows);
  const lines = summary.lines.map((l) => l.line).filter((l) => l !== "(No line)");
  // like Dev Plan: a line per event on the calendar, the processes of the line picked under it
  const events = setupEvents(rows);
  const ev = events.find((e) => e.name === line) ?? events[0] ?? null;
  const shown = ev ? ordered.filter((r) => lineOf(r) === ev.name) : [];
  const days = planDays(events);
  const evDays = ev ? calendarDays([eventSpan(ev), ...ev.tasks.map(taskSpan)], 1) : [];
  const stat = ev ? summary.lines.find((l) => l.line === ev.name) : undefined;
  const problemsOf = (r: SetupRow) => issuesBySeq(entries.filter((e) => e.setupRowId === r.id));

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    try { await work(); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const addRow = () => {
    const key = ++addKey.current;
    setAdding((a) => [...a, { key, f: newSetupRow(rows, a.map((x) => x.f), ev && ev.name !== "(No line)" ? ev.name : "") }]);
  };
  const setAdded = (key: number, f: SetupFields) => setAdding((a) => a.map((x) => (x.key === key ? { key, f } : x)));
  const dropAdded = (key: number) => setAdding((a) => a.filter((x) => x.key !== key));
  const cancelAdding = () => { setAdding([]); setAddTried(false); };
  const batch = setupBatch(adding.map((x) => x.f));
  const saveAdded = () => {
    setAddTried(true);
    if (batch.problem || !batch.rows.length) return;
    void run(async () => {
      const saved = await setupApi.import(modelName, batch.rows);
      setRows((r) => [...(r ?? []), ...saved]);
      cancelAdding();
    });
  };
  const save = (row: SetupRow, f: SetupFields) => run(async () => {
    const saved = await setupApi.save(modelName, row, f);
    setRows((r) => (r ?? []).map((x) => (x.id === row.id ? saved : x)));
    setEditing(null);
  });
  const remove = (row: SetupRow) => {
    const n = problemsOf(row).length;
    if (!window.confirm(`Delete No. ${row.no} ${row.processName}?${n ? ` Its ${n} problem${n > 1 ? "s stay" : " stays"} in Issues.` : ""}`)) return;
    void run(async () => {
      await setupApi.remove(modelName, row.id);
      setRows((r) => (r ?? []).filter((x) => x.id !== row.id));
      // the server keeps the problems as plain issues; reload them so they lose the link here too
      props.onChanged(await workingApi.entries(modelName));
    });
  };
  const importRows = (p: PastedSetup, withProblems: boolean) => run(async () => {
    const saved = await setupApi.import(modelName, p.rows.map(({ note: _note, ...f }) => f));
    setRows((r) => [...(r ?? []), ...saved]);
    setPasting(false);
    if (!withProblems) return;
    const created: Entry[] = [];
    for (let i = 0; i < saved.length; i++) {
      const note = p.rows[i].note.trim();
      if (!note) continue;
      created.push(await workingApi.addEntry(modelName, {
        ...blankEntry("issue", today, props.modelPhase), title: note.slice(0, 200), body: note.length > 200 ? note : "",
        process: issueProcessOf(saved[i]), setupRowId: saved[i].id,
      }));
    }
    if (created.length) props.onChanged([...entries, ...created]);
  });

  const toggle = (id: string) => setOpen((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const late = (date: string, state: SetupState) => !!date && date < today && state !== "O";
  const pick = (id: string) => { setLine(events.find((e) => e.id === id)?.name ?? ""); setEditing(null); };
  const byId = new Map(rows.map((r) => [r.id, r]));

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr("Plan")} <span className="cnt">{rows.length}</span></h2>
        <div className="pi-legend">
          <button className="btn sm" disabled={pasting} onClick={() => setPasting(true)}>Paste from Excel</button>
          <button className="btn sm primary" disabled={busy} onClick={addRow}>+ Add row</button>
        </div>
      </div>

      <div className="pi-sec-b">
        {pasting && <PasteBox today={today} busy={busy} onImport={importRows} onCancel={() => setPasting(false)} />}
        {!rows.length && !adding.length && (
          <div className="muted">{tr("No process yet. Add a row, or paste the rows of the plan sheet from Excel.")}</div>
        )}
        {events.length > 0 && (days.length
          ? <Gantt rows={events.map((e) => {
            const [done, all] = progress(e);
            return { id: e.id, label: e.name, sub: `${dates(eventSpan(e))} · ${done}/${all}`, span: eventSpan(e), color: e.color, done: done === all };
          })} days={days} selected={ev?.id} onPick={pick} />
          : <div className="muted">{tr("Give the processes a machine plan or setup plan day to see them on the calendar.")}</div>)}
        {events.length > 0 && !days.length && (
          <div className="wk-dp-chips">{events.map((e) => <button key={e.id} className={"wk-chip" + (e.id === ev?.id ? " on" : "")} onClick={() => pick(e.id)}>{e.name}</button>)}</div>
        )}

        {(ev || adding.length > 0) && (
          <div className={ev ? "wk-dp-detail" : undefined}>
            {ev && (
              <>
                <div className="wk-dp-dh">
                  <span className="dot" style={{ background: ev.color }} />
                  <h3>{ev.name}</h3>
                  <span className="muted small">
                    {dates(eventSpan(ev))} · {progress(ev)[0]}/{progress(ev)[1]} done
                    {stat && <> · Machine {stat.machineFinish}/{stat.total} ({rate(stat.machineFinish, stat.total)}) · Setup {stat.setupFinish}/{stat.total} ({rate(stat.setupFinish, stat.total)})</>}
                  </span>
                  <span className="wk-dp-act su-key">
                    {SETUP_STATES.map((s) => <span key={s.value}><b className={"su-st " + stateClass(s.value)}>{s.value}</b> {s.label}</span>)}
                  </span>
                </div>
                <div className="wk-qty-h static">{tr("Detailed schedule")}</div>
                {evDays.length > 0 && (
                  <Gantt days={evDays} rows={ev.tasks.map((t) => {
                    const r = byId.get(t.id);
                    return { id: t.id, label: t.name, sub: dates(taskSpan(t)) + (r?.setupStatus ? " · " + r.setupStatus : ""), span: taskSpan(t), color: ev.color, done: t.done };
                  })} />
                )}
              </>
            )}
            <div className="pi-wrap">
              <table className="wk-it su-t">
                <thead>
                  <tr>
                    <th rowSpan={2}></th><th rowSpan={2}>No</th><th rowSpan={2}>No.<br />Layout</th><th rowSpan={2}>LINE</th>
                    <th rowSpan={2} className="l">{tr("Process Name")}</th>
                    <th colSpan={3} className="grp">Machine Line</th><th colSpan={3} className="grp">SETUP</th>
                    <th rowSpan={2} className="l">{tr("Issue (Problem)")}</th><th rowSpan={2}></th>
                  </tr>
                  <tr><th>Plan</th><th>State</th><th>Note</th><th>Plan</th><th>Finish</th><th>Status</th></tr>
                </thead>
                <tbody>
                  {shown.map((r) => {
                    if (editing === r.id)
                      return <SetupEditRow key={r.id} value={setupFieldsOf(r)} busy={busy} onSave={(f) => save(r, f)} onCancel={() => setEditing(null)} />;
                    const problems = problemsOf(r);
                    const openOnes = problems.filter(isOpenIssue);
                    const isOpen = open.has(r.id);
                    return (
                      <Fragment key={r.id}>
                        <tr className={"wk-it-row su-row" + (isOpen ? " on" : "") + (r.setupStatus === "O" ? " ok" : "")} onClick={() => toggle(r.id)}>
                          <td className="tog">{isOpen ? "▾" : "▸"}</td>
                          <td className="n">{r.no}</td>
                          <td>{r.layoutNo}</td>
                          <td className="su-line-c">{r.line}</td>
                          <td className="l su-proc-c">{tr(r.processName)}</td>
                          <td className={late(r.machinePlan, r.machineState) ? "late" : ""} title={r.machinePlan}>{shortDate(r.machinePlan, today)}</td>
                          <td className={"su-st " + stateClass(r.machineState)}>{r.machineState}</td>
                          <td>{tr(r.machineNote)}</td>
                          <td className={late(r.setupPlan, r.setupStatus) ? "late" : ""} title={r.setupPlan}>{shortDate(r.setupPlan, today)}</td>
                          <td title={r.setupFinish}>{shortDate(r.setupFinish, today)}</td>
                          <td className={"su-st " + stateClass(r.setupStatus)}>{r.setupStatus}</td>
                          <td className="l su-prob">
                            {openOnes.slice(0, 3).map((p) => <div key={p.id} className="su-prob-open">- {p.title}</div>)}
                            {openOnes.length > 3 && <div className="muted small">+{openOnes.length - 3} more open</div>}
                            {problems.length > openOnes.length && (
                              <div className="muted small">{problems.length - openOnes.length} closed</div>
                            )}
                          </td>
                          <td className="act" onClick={(e) => e.stopPropagation()}>
                            <button className="wk-link" onClick={() => setEditing(r.id)}>Edit</button>
                            <button className="wk-link bad" onClick={() => remove(r)}>Delete</button>
                          </td>
                        </tr>
                        {isOpen && (
                          <tr className="su-sub">
                            <td colSpan={COLS}>
                              <IssueTable modelName={modelName} modelPhase={props.modelPhase} entries={problems} today={today}
                                setupRowId={r.id} heading={`Issue (Problem) · No. ${r.no} ${r.processName}`} newDefaults={{ process: issueProcessOf(r) }}
                                onChanged={(list) => {
                                  const mine = new Set(problems.map((p) => p.id));
                                  props.onChanged([...entries.filter((e) => !mine.has(e.id)), ...list]);
                                }}
                                onError={props.onError} />
                            </td>
                          </tr>
                        )}
                      </Fragment>
                    );
                  })}
                  {adding.map((x, i) => {
                    const problem = addTried ? setupProblem(x.f) : "";
                    return (
                      <tr key={x.key} className="su-edit su-new"
                        onKeyDown={(e) => { if (e.key === "Enter" && e.target instanceof HTMLInputElement) saveAdded(); }}>
                        <td className="tog">+</td>
                        <SetupCells value={x.f} onChange={(f) => setAdded(x.key, f)} autoFocus={i === adding.length - 1} />
                        <td className="l su-msg">{problem && <span className="su-err">{problem}</span>}</td>
                        <td className="act"><button className="wk-link bad" title="Remove this new row" onClick={() => dropAdded(x.key)}>Remove</button></td>
                      </tr>
                    );
                  })}
                  {adding.length > 0 && (
                    <tr className="su-add-f">
                      <td colSpan={COLS}>
                        <button className="btn sm" disabled={busy} onClick={addRow}>+ Another row</button>
                        <button className="btn sm primary" disabled={busy} onClick={saveAdded}>Add {adding.length} row{adding.length > 1 ? "s" : ""}</button>
                        <button className="btn sm" disabled={busy} onClick={cancelAdding}>Cancel</button>
                        {addTried && batch.problem ? <span className="su-err">{batch.problem}</span>
                          : <span className="muted small">Enter adds them all. Problems are added once a row is saved.</span>}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
              <datalist id="su-lines">{lines.map((l) => <option key={l} value={l} />)}</datalist>
            </div>
          </div>
        )}
      </div>
    </section>
  );
}
