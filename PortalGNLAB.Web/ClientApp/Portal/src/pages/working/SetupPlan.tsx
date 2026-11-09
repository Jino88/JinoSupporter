import { Fragment, useEffect, useState } from "react";
import { workingApi } from "./api";
import IssueTable from "./IssueTable";
import { blankEntry, isOpenIssue, issuesBySeq } from "./logic";
import {
  blankSetup, nextNo, parseSetupPaste, rate, rowsInOrder, SETUP_STATES, setupApi, setupFieldsOf, setupProblem, setupSummary, shortDate,
} from "./setup";
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
      {SETUP_STATES.map((s) => <option key={s.value} value={s.value}>{s.value} {s.label}</option>)}
    </select>
  );
}

/** A row being added or edited, in place, in the same columns. */
function SetupEditRow(props: {
  value: SetupFields;
  lines: string[];
  busy: boolean;
  isNew: boolean;
  onSave: (f: SetupFields) => void;
  onCancel: () => void;
}) {
  const [f, setF] = useState(props.value);
  const [shown, setShown] = useState(false);
  const set = <K extends keyof SetupFields>(k: K, v: SetupFields[K]) => setF((x) => ({ ...x, [k]: v }));
  const problem = setupProblem(f);
  const save = () => { setShown(true); if (!problem) props.onSave({ ...f, processName: f.processName.trim() }); };
  const date = (k: "machinePlan" | "setupPlan" | "setupFinish", label: string) => (
    <td><input className="inp sm su-date" type="date" aria-label={label} value={f[k]} onChange={(e) => set(k, e.target.value)} /></td>
  );
  return (
    <tr className="su-edit" onKeyDown={(e) => { if (e.key === "Escape") props.onCancel(); if (e.key === "Enter" && e.target instanceof HTMLInputElement) save(); }}>
      <td></td>
      <td><input className="inp sm su-no" type="number" min={1} autoFocus aria-label="No." value={f.no || ""}
        onChange={(e) => set("no", e.target.value === "" ? 0 : Number(e.target.value))} /></td>
      <td><input className="inp sm su-no" aria-label="No. layout" value={f.layoutNo} onChange={(e) => set("layoutNo", e.target.value)} /></td>
      <td>
        <input className="inp sm su-line" aria-label="Line" list="su-lines" value={f.line} onChange={(e) => set("line", e.target.value)} />
        <datalist id="su-lines">{props.lines.map((l) => <option key={l} value={l} />)}</datalist>
      </td>
      <td className="l"><input className="inp sm su-proc" aria-label="Process name" value={f.processName} onChange={(e) => set("processName", e.target.value)} /></td>
      {date("machinePlan", "Machine plan")}
      <td><StateSelect label="Machine state" value={f.machineState} onChange={(v) => set("machineState", v)} /></td>
      <td><input className="inp sm su-note" aria-label="Machine note" value={f.machineNote} onChange={(e) => set("machineNote", e.target.value)} /></td>
      {date("setupPlan", "Setup plan")}
      {date("setupFinish", "Setup finish")}
      <td><StateSelect label="Setup status" value={f.setupStatus} onChange={(v) => set("setupStatus", v)} /></td>
      <td className="l">
        {shown && problem ? <span className="su-err">{problem}</span>
          : <span className="muted small">{props.isNew ? "Add problems after saving the row." : "Problems are edited below the row."}</span>}
      </td>
      <td className="act">
        <button className="btn sm primary" disabled={props.busy} onClick={save}>{props.isNew ? "Add" : "Save"}</button>
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
        onChange={(e) => setText(e.target.value)} />
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
 * Setup Plan: per process, the machine line plan / state and the setup plan / finish / status, summed per line at the top.
 * Each process opens into its Issue (Problem) list: issues of this model raised against the row, with the same columns,
 * daily notes and pictures as in Issues (they are listed there too).
 */
export default function SetupPlan(props: {
  modelName: string;
  modelPhase: string;
  entries: Entry[];
  today: string;
  onChanged: (list: Entry[]) => void;
  onError: (e: unknown) => void;
}) {
  const { modelName, entries, today } = props;
  const [rows, setRows] = useState<SetupRow[] | null>(null);
  const [editing, setEditing] = useState<string | null>(null); // row id, "new" or null
  const [pasting, setPasting] = useState(false);
  const [open, setOpen] = useState<Set<string>>(new Set());
  const [line, setLine] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let live = true;
    setupApi.rows(modelName).then((r) => { if (live) setRows(r); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [modelName]);

  if (rows === null) return <div className="muted">Loading the setup plan…</div>;

  const ordered = rowsInOrder(rows);
  const summary = setupSummary(rows);
  const lines = summary.lines.map((l) => l.line).filter((l) => l !== "(No line)");
  const shown = line ? ordered.filter((r) => (r.line.trim() || "(No line)") === line) : ordered;
  const problemsOf = (r: SetupRow) => issuesBySeq(entries.filter((e) => e.setupRowId === r.id));

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    try { await work(); setEditing(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const add = (f: SetupFields) => run(async () => { const saved = await setupApi.add(modelName, f); setRows((r) => [...(r ?? []), saved]); });
  const save = (row: SetupRow, f: SetupFields) => run(async () => {
    const saved = await setupApi.save(modelName, row, f);
    setRows((r) => (r ?? []).map((x) => (x.id === row.id ? saved : x)));
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
        process: saved[i].line, setupRowId: saved[i].id,
      }));
    }
    if (created.length) props.onChanged([...entries, ...created]);
  });

  const toggle = (id: string) => setOpen((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const late = (date: string, state: SetupState) => !!date && date < today && state !== "O";

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>■ {modelName} Setup Plan <span className="cnt">{rows.length}</span></h2>
        <div className="pi-legend">
          <button className="btn sm" disabled={pasting} onClick={() => setPasting(true)}>Paste from Excel</button>
          <button className="btn sm primary" disabled={editing === "new"} onClick={() => setEditing("new")}>+ Add row</button>
        </div>
      </div>

      <div className="pi-sec-b su-top">
        <table className="su-sum">
          <thead>
            <tr><th rowSpan={2}>TYPE</th><th rowSpan={2}>Total<br />Process</th><th colSpan={3}>Machine Status</th><th colSpan={2}>Setup Status</th></tr>
            <tr><th>Total</th><th>Finish</th><th>Rate</th><th>Finish</th><th>Rate</th></tr>
          </thead>
          <tbody>
            {summary.lines.map((s) => (
              <tr key={s.line} className={line === s.line ? "on" : ""} onClick={() => setLine(line === s.line ? "" : s.line)} title="Show only this line">
                <th>{s.line}</th><td>{s.total}</td><td className="v">{s.total}</td><td className="v">{s.machineFinish}</td>
                <td className="v">{rate(s.machineFinish, s.total)}</td><td className="v">{s.setupFinish}</td><td className="v">{rate(s.setupFinish, s.total)}</td>
              </tr>
            ))}
            <tr className="tot">
              <th>Total</th><td>{summary.total.total}</td><td>{summary.total.total}</td><td>{summary.total.machineFinish}</td>
              <td>{rate(summary.total.machineFinish, summary.total.total)}</td><td>{summary.total.setupFinish}</td>
              <td>{rate(summary.total.setupFinish, summary.total.total)}</td>
            </tr>
          </tbody>
        </table>
        <table className="su-sum su-legend">
          <thead><tr><th colSpan={2}>Setup Status</th></tr></thead>
          <tbody>{SETUP_STATES.map((s) => <tr key={s.value}><td className={"su-st " + stateClass(s.value)}>{s.value}</td><td>{s.label}</td></tr>)}</tbody>
        </table>
        {line && <button className="wk-chip on su-filter" onClick={() => setLine("")}>Line: {line} ×</button>}
      </div>

      {pasting && (
        <div className="pi-sec-b">
          <PasteBox today={today} busy={busy} onImport={importRows} onCancel={() => setPasting(false)} />
        </div>
      )}

      <div className="pi-wrap">
        <table className="wk-it su-t">
          <thead>
            <tr>
              <th rowSpan={2}></th><th rowSpan={2}>No</th><th rowSpan={2}>No.<br />Layout</th><th rowSpan={2}>LINE</th>
              <th rowSpan={2} className="l">Process Name</th>
              <th colSpan={3} className="grp">Machine Line</th><th colSpan={3} className="grp">SETUP</th>
              <th rowSpan={2} className="l">Issue (Problem)</th><th rowSpan={2}></th>
            </tr>
            <tr><th>Plan</th><th>State</th><th>Note</th><th>Plan</th><th>Finish</th><th>Status</th></tr>
          </thead>
          <tbody>
            {shown.map((r, i) => {
              if (editing === r.id)
                return <SetupEditRow key={r.id} value={setupFieldsOf(r)} lines={lines} busy={busy} isNew={false}
                  onSave={(f) => save(r, f)} onCancel={() => setEditing(null)} />;
              const problems = problemsOf(r);
              const openOnes = problems.filter(isOpenIssue);
              const isOpen = open.has(r.id);
              const sameLine = i > 0 && shown[i - 1].line === r.line;
              return (
                <Fragment key={r.id}>
                  <tr className={"wk-it-row su-row" + (isOpen ? " on" : "") + (r.setupStatus === "O" ? " ok" : "")} onClick={() => toggle(r.id)}>
                    <td className="tog">{isOpen ? "▾" : "▸"}</td>
                    <td className="n">{r.no}</td>
                    <td>{r.layoutNo}</td>
                    <td className={"su-line-c" + (sameLine ? " same" : "")}>{r.line}</td>
                    <td className="l su-proc-c">{r.processName}</td>
                    <td className={late(r.machinePlan, r.machineState) ? "late" : ""} title={r.machinePlan}>{shortDate(r.machinePlan, today)}</td>
                    <td className={"su-st " + stateClass(r.machineState)}>{r.machineState}</td>
                    <td>{r.machineNote}</td>
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
                          setupRowId={r.id} heading={`Issue (Problem) · No. ${r.no} ${r.processName}`} newDefaults={{ process: r.line }}
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
            {editing === "new" && (
              <SetupEditRow value={{ ...blankSetup(line || ordered.at(-1)?.line || ""), no: nextNo(rows) }} lines={lines} busy={busy} isNew
                onSave={add} onCancel={() => setEditing(null)} />
            )}
            {!rows.length && editing !== "new" && (
              <tr><td colSpan={COLS} className="wk-it-empty">No process yet. Add a row, or paste the rows of the Setup Plan sheet from Excel.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}
