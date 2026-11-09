import { Fragment, useState } from "react";
import { workingApi } from "./api";
import { blankEntry, entryProblem, fieldsOf, isLate, isOpenIssue, issueFields, issuesBySeq, logProblem, suggestions } from "./logic";
import { ISSUE_CAUSES, ISSUE_PROCESSES, ISSUE_TYPES, PHASES, type Entry, type EntryFields, type EntryLog } from "./types";

type Filter = "open" | "closed" | "all";
const COLS = 14;

interface Lists { types: string[]; processes: string[]; causes: string[] }

function IssueForm(props: { value: EntryFields; lists: Lists; busy: boolean; onSave: (f: EntryFields) => void; onCancel: () => void }) {
  const [f, setF] = useState(props.value);
  const set = <K extends keyof EntryFields>(k: K, v: EntryFields[K]) => setF((cur) => ({ ...cur, [k]: v }));
  const problem = entryProblem(f);
  const text = (k: "category" | "process" | "cause" | "owner", label: string, list?: string) => (
    <div className="wk-f"><label>{label}</label>
      <input className="inp" value={f[k]} list={list} onChange={(e) => set(k, e.target.value)} /></div>
  );
  const date = (k: "date" | "due" | "doneDate", label: string) => (
    <div className="wk-f"><label>{label}</label>
      <input type="date" className="inp" value={f[k]} onChange={(e) => set(k, e.target.value)} /></div>
  );
  return (
    <div className="wk-form wk-ef wk-if">
      <div className="wk-f"><label>Dev phase</label>
        <select className="inp" value={f.phase} onChange={(e) => set("phase", e.target.value)}>
          <option value="">—</option>
          {PHASES.map((p) => <option key={p}>{p}</option>)}
          {f.phase && !PHASES.includes(f.phase as (typeof PHASES)[number]) && <option>{f.phase}</option>}
        </select>
      </div>
      {text("category", "Issue type", "wk-l-type")}
      {text("process", "Process", "wk-l-proc")}
      {text("cause", "Cause group", "wk-l-cause")}
      <div className="wk-f wide"><label>Issue</label>
        <input className="inp" value={f.title} maxLength={200} autoFocus onChange={(e) => set("title", e.target.value)} /></div>
      <div className="wk-f wide"><label>Details</label>
        <textarea className="inp wk-ta" rows={3} value={f.body} onChange={(e) => set("body", e.target.value)} /></div>
      {text("owner", "Owner")}
      {date("date", "Registered")}
      {date("due", "Due")}
      {date("doneDate", "Done")}
      <div className="wk-f wide"><label>Final result</label>
        <textarea className="inp wk-ta" rows={2} value={f.finalResult} onChange={(e) => set("finalResult", e.target.value)} /></div>
      <div className="wk-f wide wk-ef-act">
        {problem && f.title && <span className="pi-dirty">{problem}</span>}
        <button className="btn sm" onClick={props.onCancel} disabled={props.busy}>Cancel</button>
        <button className="btn sm primary" disabled={!!problem || props.busy} onClick={() => props.onSave(issueFields(f))}>{props.busy ? "Saving…" : "Save"}</button>
      </div>
      <datalist id="wk-l-type">{props.lists.types.map((v) => <option key={v} value={v} />)}</datalist>
      <datalist id="wk-l-proc">{props.lists.processes.map((v) => <option key={v} value={v} />)}</datalist>
      <datalist id="wk-l-cause">{props.lists.causes.map((v) => <option key={v} value={v} />)}</datalist>
    </div>
  );
}

/** Daily progress notes under one issue: oldest first, with a box at the bottom for today's note. */
function Notes(props: { modelName: string; issue: Entry; today: string; onLogs: (logs: EntryLog[]) => void; onError: (e: unknown) => void }) {
  const { modelName, issue, today } = props;
  const [date, setDate] = useState(today);
  const [text, setText] = useState("");
  const [edit, setEdit] = useState<{ id: string; date: string; text: string } | null>(null);
  const [busy, setBusy] = useState(false);
  const logs = issue.logs ?? [];

  const run = async (work: () => Promise<EntryLog[]>) => {
    setBusy(true);
    try { props.onLogs(await work()); return true; } catch (e) { props.onError(e); return false; } finally { setBusy(false); }
  };
  const add = async () => {
    const ok = await run(async () => [...logs, await workingApi.addLog(modelName, issue.id, date, text)]);
    if (ok) setText("");
  };
  const save = async (log: EntryLog) => {
    if (!edit) return;
    const ok = await run(async () => {
      const saved = await workingApi.saveLog(modelName, issue.id, log, edit.date, edit.text);
      return logs.map((l) => (l.id === log.id ? saved : l));
    });
    if (ok) setEdit(null);
  };
  const remove = (log: EntryLog) => {
    if (!window.confirm(`Delete the note of ${log.date}?`)) return;
    void run(async () => { await workingApi.deleteLog(modelName, issue.id, log.id); return logs.filter((l) => l.id !== log.id); });
  };
  const sorted = [...logs].sort((a, b) => a.date.localeCompare(b.date) || a.createdAt.localeCompare(b.createdAt));
  const addProblem = logProblem(date, text);

  return (
    <div className="wk-notes">
      {issue.body && <div className="wk-notes-body"><b>Details</b><div className="wk-text">{issue.body}</div></div>}
      <div className="wk-notes-h">Daily progress <span className="cnt">{logs.length}</span></div>
      {!sorted.length && <div className="muted small">No progress note yet.</div>}
      {sorted.map((l) => (
        <div key={l.id} className="wk-note">
          {edit?.id === l.id ? (
            <>
              <input type="date" className="inp sm" value={edit.date} onChange={(e) => setEdit({ ...edit, date: e.target.value })} />
              <textarea className="inp wk-ta" rows={2} value={edit.text} onChange={(e) => setEdit({ ...edit, text: e.target.value })} />
              <div className="wk-note-act">
                <button className="btn sm" onClick={() => setEdit(null)} disabled={busy}>Cancel</button>
                <button className="btn sm primary" onClick={() => save(l)} disabled={busy || !!logProblem(edit.date, edit.text)}>Save</button>
              </div>
            </>
          ) : (
            <>
              <span className="wk-note-d">{l.date}</span>
              <div className="wk-text">{l.text}</div>
              <div className="wk-note-act">
                <span className="muted">{l.createdBy || "—"}</span>
                <button className="wk-link" onClick={() => setEdit({ id: l.id, date: l.date, text: l.text })} disabled={busy}>Edit</button>
                <button className="wk-link danger" onClick={() => remove(l)} disabled={busy}>Delete</button>
              </div>
            </>
          )}
        </div>
      ))}
      <div className="wk-note wk-note-new">
        <input type="date" className="inp sm" value={date} onChange={(e) => setDate(e.target.value)} aria-label="Note date" />
        <textarea className="inp wk-ta" rows={2} value={text} placeholder="What was done today…" onChange={(e) => setText(e.target.value)} />
        <div className="wk-note-act">
          {text && addProblem && <span className="pi-dirty">{addProblem}</span>}
          <button className="btn sm primary" onClick={add} disabled={busy || !!addProblem}>Add note</button>
        </div>
      </div>
    </div>
  );
}

/**
 * Issue register: No., dev phase, classification (issue type / process / cause group), issue, owner,
 * registered / due / done dates and final result, with daily progress notes under each issue.
 */
export default function IssueTable(props: {
  modelName: string;
  modelPhase: string;
  entries: Entry[];
  today: string;
  onChanged: (list: Entry[]) => void;
  onError: (e: unknown) => void;
}) {
  const { modelName, entries, today } = props;
  const [filter, setFilter] = useState<Filter>("open");
  const [editing, setEditing] = useState<string | null>(null); // issue id, "new" or null
  const [open, setOpen] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);

  const all = issuesBySeq(entries);
  const shown = all.filter((e) => filter === "all" || (filter === "open") === isOpenIssue(e));
  const lists: Lists = {
    types: suggestions(all.map((e) => e.category), ISSUE_TYPES),
    processes: suggestions(all.map((e) => e.process), ISSUE_PROCESSES),
    causes: suggestions(all.map((e) => e.cause), ISSUE_CAUSES),
  };

  const run = async (work: () => Promise<Entry[]>) => {
    setBusy(true);
    try { props.onChanged(await work()); setEditing(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const add = (f: EntryFields) => run(async () => [...entries, await workingApi.addEntry(modelName, f)]);
  const save = (e: Entry, f: EntryFields) =>
    run(async () => { const saved = await workingApi.saveEntry(modelName, e.id, f, e.revision); return entries.map((x) => (x.id === e.id ? saved : x)); });
  const remove = (e: Entry) => {
    if (!window.confirm(`Delete issue No. ${e.seq} "${e.title}" and its ${e.logs?.length ?? 0} notes?`)) return;
    void run(async () => { await workingApi.deleteEntry(modelName, e.id); return entries.filter((x) => x.id !== e.id); });
  };
  const setLogs = (e: Entry, logs: EntryLog[]) => props.onChanged(entries.map((x) => (x.id === e.id ? { ...x, logs } : x)));
  const toggle = (id: string) => setOpen((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const count = (f: Filter) => all.filter((e) => f === "all" || (f === "open") === isOpenIssue(e)).length;

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Issues <span className="cnt">{shown.length}</span></h2>
        <div className="pi-legend">
          {(["open", "closed", "all"] as Filter[]).map((v) => (
            <button key={v} className={"wk-chip" + (filter === v ? " on" : "")} onClick={() => setFilter(v)}>
              {v[0].toUpperCase() + v.slice(1)} {count(v)}
            </button>
          ))}
          <button className="btn sm primary" disabled={editing === "new"} onClick={() => setEditing("new")}>+ New issue</button>
        </div>
      </div>
      {editing === "new" && (
        <div className="pi-sec-b">
          <IssueForm value={blankEntry("issue", today, props.modelPhase)} lists={lists} busy={busy} onSave={add} onCancel={() => setEditing(null)} />
        </div>
      )}
      <div className="pi-wrap">
        <table className="wk-it">
          <thead>
            <tr>
              <th rowSpan={2}></th><th rowSpan={2}>No.</th><th rowSpan={2}>Dev phase</th>
              <th colSpan={3} className="grp">Classification</th>
              <th rowSpan={2} className="l">Issue</th><th rowSpan={2}>Owner</th>
              <th rowSpan={2}>Registered</th><th rowSpan={2}>Due</th><th rowSpan={2}>Done</th>
              <th rowSpan={2} className="l">Final result</th><th rowSpan={2}>Notes</th><th rowSpan={2}></th>
            </tr>
            <tr><th>Issue type</th><th>Process</th><th>Cause group</th></tr>
          </thead>
          <tbody>
            {!shown.length && (
              <tr><td colSpan={COLS} className="wk-it-empty">{all.length ? "No issue in this filter." : "No issue yet. Add one with + New issue."}</td></tr>
            )}
            {shown.map((e) => {
              const expanded = open.has(e.id);
              const last = e.logs?.length ? [...e.logs].sort((a, b) => b.date.localeCompare(a.date))[0] : null;
              return (
                <Fragment key={e.id}>
                  <tr className={"wk-it-row" + (isOpenIssue(e) ? "" : " done") + (expanded ? " on" : "")} onClick={() => toggle(e.id)}>
                    <td className="tog">{expanded ? "▾" : "▸"}</td>
                    <td className="n">{e.seq || "—"}</td>
                    <td>{e.phase || "—"}</td>
                    <td>{e.category}</td>
                    <td>{e.process}</td>
                    <td>{e.cause}</td>
                    <td className="l wk-it-title"><b>{e.title}</b>{last && <div className="muted small">{last.date.slice(5)} · {last.text.split("\n")[0]}</div>}</td>
                    <td>{e.owner}</td>
                    <td>{e.date}</td>
                    <td className={isLate(e, today) ? "late" : ""}>{e.due}{isLate(e, today) && <span className="wk-late">Late</span>}</td>
                    <td>{e.doneDate ? <span className="wk-st closed">{e.doneDate}</span> : <span className="wk-st open">Open</span>}</td>
                    <td className="l wk-it-final">{e.finalResult}</td>
                    <td className="n">{e.logs?.length ?? 0}</td>
                    <td className="act" onClick={(ev) => ev.stopPropagation()}>
                      <button className="wk-link" onClick={() => setEditing(e.id)} disabled={busy}>Edit</button>
                      <button className="wk-link danger" onClick={() => remove(e)} disabled={busy}>Delete</button>
                    </td>
                  </tr>
                  {editing === e.id && (
                    <tr className="wk-it-sub"><td colSpan={COLS}>
                      <IssueForm value={fieldsOf(e)} lists={lists} busy={busy} onSave={(f) => save(e, f)} onCancel={() => setEditing(null)} />
                    </td></tr>
                  )}
                  {expanded && editing !== e.id && (
                    <tr className="wk-it-sub"><td colSpan={COLS}>
                      <Notes modelName={modelName} issue={e} today={today} onLogs={(logs) => setLogs(e, logs)} onError={props.onError} />
                    </td></tr>
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}
