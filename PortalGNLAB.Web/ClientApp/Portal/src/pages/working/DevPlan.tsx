import { useEffect, useMemo, useState } from "react";
import { WEEKDAYS, localToday, weekday } from "../productionInput/logic";
import { workingApi } from "./api";
import {
  COLORS, blankEvent, blankRow, blankTask, calendarDays, eventSpan, inSpan, monthRuns, planProblem, progress, qtyDays, readDevPlan,
  readQty, rowTotal, setQty, taskSpan,
} from "./plan";
import type { DevPlan as Plan, PlanEvent, PlanTask, QtyRow } from "./plan";
import type { WorkingModel } from "./types";

interface Row { id: string; label: string; sub: string; span: [string, string] | null; color: string; done?: boolean }

const dates = (span: [string, string] | null) => (span ? (span[0] === span[1] ? span[0].slice(5) : `${span[0].slice(5)} – ${span[1].slice(5)}`) : "no date");

/** Horizontal calendar: a column per day, a row per item, the days of each item filled in its color. */
function Gantt(props: { rows: Row[]; days: string[]; selected?: string; onPick?: (id: string) => void }) {
  const today = localToday();
  const { rows, days } = props;
  return (
    <div className="wk-gantt-wrap">
      {/* a short calendar fills the width; a long one scrolls */}
      <table className={"wk-gantt" + (days.length <= 60 ? " fill" : "")}>
        <thead>
          <tr>
            <th className="nm" rowSpan={2}></th>
            {monthRuns(days).map(([m, n]) => <th key={m} colSpan={n} className="mon">{m}</th>)}
          </tr>
          <tr>
            {days.map((d) => (
              <th key={d} className={"day" + (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "")} title={d}>
                {Number(d.slice(8))}<small>{WEEKDAYS[weekday(d)][0]}</small>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.id} className={(props.onPick ? "pick" : "") + (r.id === props.selected ? " on" : "")}
              onClick={props.onPick ? () => props.onPick!(r.id) : undefined}>
              <td className="nm">
                <span className="dot" style={{ background: r.color }} />
                <b className={r.done ? "done" : ""}>{r.label}</b>
                <small>{r.sub}</small>
              </td>
              {days.map((d) => {
                const on = inSpan(d, r.span);
                const cls = "c" + (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "")
                  + (on ? " on" + (d === r.span![0] ? " s" : "") + (d === r.span![1] ? " e" : "") : "");
                return <td key={d} className={cls} style={on ? { background: r.done ? "#94a3b8" : r.color } : undefined} />;
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

const num = (n: number) => n.toLocaleString("en-US", { maximumFractionDigits: 3 });

/** One day of a quantity row: typed freely, kept when it leaves the cell (Enter or Tab), put back when not a number. */
function QtyCell(props: { value: number | undefined; label: string; onKeep: (n: number | undefined) => void; onPaste: (text: string) => boolean }) {
  const shown = props.value ? String(props.value) : "";
  const [text, setText] = useState(shown);
  useEffect(() => setText(shown), [shown]);
  const keep = () => {
    const n = readQty(text);
    if (n === null) setText(shown);
    else if (n !== props.value) props.onKeep(n);
  };
  return (
    <input className="wk-qty-in" value={text} inputMode="decimal" aria-label={props.label}
      onChange={(x) => setText(x.target.value)} onBlur={keep}
      onKeyDown={(x) => { if (x.key === "Enter") x.currentTarget.blur(); if (x.key === "Escape") { setText(shown); } }}
      onPaste={(x) => { if (props.onPaste(x.clipboardData.getData("text"))) x.preventDefault(); }} />
  );
}

/**
 * Quantity table of an event, as in the build plan sheet: a row per stage (SPK Assy, Lower Assy, Function, OQC, ...)
 * with its total, and a column per day to type the quantity of that day. Rows can be added, renamed, moved and deleted;
 * cells copied from Excel can be pasted from the cell clicked.
 */
function QtyTable(props: { ev: PlanEvent; onRows: (rows: QtyRow[]) => void }) {
  const { ev } = props;
  const today = localToday();
  const days = useMemo(() => qtyDays(ev), [ev]);
  const span = eventSpan(ev);
  const rows = ev.rows;
  const setRow = (r: QtyRow) => props.onRows(rows.map((x) => (x.id === r.id ? r : x)));
  const move = (i: number, by: -1 | 1) => {
    const j = i + by;
    if (j < 0 || j >= rows.length) return;
    const list = [...rows];
    [list[i], list[j]] = [list[j], list[i]];
    props.onRows(list);
  };
  const remove = (r: QtyRow) => {
    if (Object.keys(r.qty).length && !window.confirm(`Delete the row ${r.name || "(no name)"} and its quantities?`)) return;
    props.onRows(rows.filter((x) => x.id !== r.id));
  };
  /** Cells copied from Excel (more than one) go to the right and down from the cell; numbers only, blanks clear. */
  const paste = (ri: number, di: number, text: string) => {
    const grid = text.replace(/\r?\n$/, "").split(/\r?\n/).map((l) => l.split("\t"));
    if (grid.length === 1 && grid[0].length === 1) return false;
    const list = [...rows];
    grid.forEach((line, y) => {
      const r = list[ri + y];
      if (!r) return;
      let next = r;
      line.forEach((cell, x) => {
        const d = days[di + x];
        const n = readQty(cell);
        if (d && n !== null) next = setQty(next, d, n);
      });
      list[ri + y] = next;
    });
    props.onRows(list);
    return true;
  };
  const dayCls = (d: string) => (d === today ? " today" : "") + (weekday(d) === 6 ? " sat" : weekday(d) === 0 ? " sun" : "") + (inSpan(d, span) ? " in" : "");
  return (
    <div className="wk-qty">
      {days.length ? (
        <div className="wk-qty-wrap">
          <table className="wk-qty-t">
            <thead>
              <tr>
                <th className="nm">Item</th>
                <th className="tot">Total</th>
                {days.map((d) => (
                  <th key={d} className={"day" + dayCls(d)} title={d} style={inSpan(d, span) ? { boxShadow: `inset 0 -3px 0 ${ev.color}` } : undefined}>
                    {Number(d.slice(5, 7))}/{Number(d.slice(8))}<small>{WEEKDAYS[weekday(d)]}</small>
                  </th>
                ))}
                <th className="act"></th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, ri) => (
                <tr key={r.id}>
                  <td className="nm">
                    <input className="wk-qty-name" value={r.name} placeholder="Item, e.g. SPK Assy" aria-label="Item name"
                      onChange={(x) => setRow({ ...r, name: x.target.value })} />
                  </td>
                  <td className="tot">{rowTotal(r) ? num(rowTotal(r)) : ""}</td>
                  {days.map((d, di) => (
                    <td key={d} className={"c" + dayCls(d) + (r.qty[d] ? " has" : "")}>
                      <QtyCell value={r.qty[d]} label={`${r.name} ${d}`} onKeep={(n) => setRow(setQty(r, d, n))} onPaste={(t) => paste(ri, di, t)} />
                    </td>
                  ))}
                  <td className="act">
                    <button className="wk-link" onClick={() => move(ri, -1)} disabled={ri === 0} aria-label="Move up">↑</button>
                    <button className="wk-link" onClick={() => move(ri, 1)} disabled={ri === rows.length - 1} aria-label="Move down">↓</button>
                    <button className="wk-link bad" onClick={() => remove(r)} aria-label={`Delete ${r.name}`}>×</button>
                  </td>
                </tr>
              ))}
              {!rows.length && <tr><td colSpan={days.length + 3} className="wk-it-empty">No item yet. Add the items to plan (SPK Assy, Lower Assy, …) with + Row.</td></tr>}
            </tbody>
          </table>
        </div>
      ) : <div className="muted small">Give the event dates to plan quantities by day.</div>}
      <button className="btn sm wk-dp-add" onClick={() => props.onRows([...rows, blankRow()])}>+ Row</button>
    </div>
  );
}

/** Name, dates and color of an event being added or changed. */
function EventEdit(props: { value: PlanEvent; onKeep: (e: PlanEvent) => void; onCancel: () => void }) {
  const [e, setE] = useState(props.value);
  const problem = planProblem(e);
  return (
    <div className="wk-dp-form">
      <input className="inp sm" value={e.name} placeholder="Event, e.g. DR build" autoFocus aria-label="Event name"
        onChange={(x) => setE({ ...e, name: x.target.value })} />
      <input className="inp sm" type="date" value={e.start} aria-label="Start" onChange={(x) => setE({ ...e, start: x.target.value })} />
      <span className="muted">–</span>
      <input className="inp sm" type="date" value={e.end} aria-label="End" onChange={(x) => setE({ ...e, end: x.target.value })} />
      <span className="wk-dp-colors">
        {COLORS.map((c) => (
          <button key={c} className={"sw" + (c === e.color ? " on" : "")} style={{ background: c }} aria-label={"Color " + c}
            onClick={() => setE({ ...e, color: c })} />
        ))}
      </span>
      <span className="muted small">No dates = from its tasks.</span>
      {problem && e.name.trim() && <span className="pi-dirty small">{problem}</span>}
      <button className="btn sm" onClick={props.onCancel}>Cancel</button>
      <button className="btn sm primary" disabled={!!problem} onClick={() => props.onKeep({ ...e, name: e.name.trim() })}>Keep</button>
    </div>
  );
}

/** One task of the detailed schedule being added or changed. */
function TaskEdit(props: { value: PlanTask; onKeep: (t: PlanTask) => void; onCancel: () => void }) {
  const [t, setT] = useState(props.value);
  const problem = planProblem(t);
  const set = (f: Partial<PlanTask>) => setT({ ...t, ...f });
  return (
    <tr className="wk-dp-edit">
      <td><input type="checkbox" checked={t.done} aria-label="Done" onChange={(x) => set({ done: x.target.checked })} /></td>
      <td className="l"><input className="inp sm" value={t.name} placeholder="Task" autoFocus onChange={(x) => set({ name: x.target.value })} /></td>
      <td><input className="inp sm" type="date" value={t.start} aria-label="Start" onChange={(x) => set({ start: x.target.value })} /></td>
      <td><input className="inp sm" type="date" value={t.end} aria-label="End" onChange={(x) => set({ end: x.target.value })} /></td>
      <td><input className="inp sm wk-dp-owner" value={t.owner} placeholder="Owner" onChange={(x) => set({ owner: x.target.value })} /></td>
      <td className="l"><input className="inp sm" value={t.note} placeholder="Note" onChange={(x) => set({ note: x.target.value })} /></td>
      <td className="act">
        {problem && t.name.trim() && <div className="pi-dirty small">{problem}</div>}
        <button className="btn sm" onClick={props.onCancel}>Cancel</button>{" "}
        <button className="btn sm primary" disabled={!!problem} onClick={() => props.onKeep({ ...t, name: t.name.trim(), owner: t.owner.trim(), note: t.note.trim() })}>Keep</button>
      </td>
    </tr>
  );
}

/**
 * Development plan of the model: the events on a horizontal calendar (a row per event, its days filled in), and under
 * it the detailed schedule of the event clicked. Kept in the model information (info.devPlan) and saved here.
 */
export default function DevPlan(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
}) {
  const { model } = props;
  const savedPlan = useMemo(() => readDevPlan(model.info.devPlan), [model]);
  const [plan, setPlan] = useState<Plan>(savedPlan);
  const [selected, setSelected] = useState(savedPlan.events[0]?.id ?? "");
  const [editEvent, setEditEvent] = useState<PlanEvent | null>(null);
  const [editTask, setEditTask] = useState<PlanTask | null>(null);
  const [showQty, setShowQty] = useState(true);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const dirty = JSON.stringify(plan) !== JSON.stringify(savedPlan);
  useEffect(() => props.onDirty(dirty), [dirty]);

  const events = plan.events;
  const ev = events.find((e) => e.id === selected) ?? null;
  const days = useMemo(() => calendarDays(events.map(eventSpan)), [events]);
  const evDays = useMemo(() => (ev ? calendarDays([eventSpan(ev), ...ev.tasks.map(taskSpan)], 1) : []), [ev]);

  const setEvents = (list: PlanEvent[]) => setPlan({ ...plan, events: list });
  const keepEvent = (e: PlanEvent) => {
    setEvents(events.some((x) => x.id === e.id) ? events.map((x) => (x.id === e.id ? e : x)) : [...events, e]);
    setSelected(e.id);
    setEditEvent(null);
  };
  const removeEvent = (e: PlanEvent) => {
    if (!window.confirm(`Delete the event ${e.name} and its ${e.tasks.length} task(s)?`)) return;
    setEvents(events.filter((x) => x.id !== e.id));
    setSelected(events.find((x) => x.id !== e.id)?.id ?? "");
  };
  const move = (e: PlanEvent, by: -1 | 1) => {
    const i = events.indexOf(e), j = i + by;
    if (j < 0 || j >= events.length) return;
    const list = [...events];
    [list[i], list[j]] = [list[j], list[i]];
    setEvents(list);
  };
  const setRows = (rows: QtyRow[]) => ev && setEvents(events.map((x) => (x.id === ev.id ? { ...x, rows } : x)));
  const setTasks = (tasks: PlanTask[]) => ev && setEvents(events.map((x) => (x.id === ev.id ? { ...x, tasks } : x)));
  const keepTask = (t: PlanTask) => {
    if (!ev) return;
    setTasks(ev.tasks.some((x) => x.id === t.id) ? ev.tasks.map((x) => (x.id === t.id ? t : x)) : [...ev.tasks, t]);
    setEditTask(null);
  };
  const save = async () => {
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, { ...model.info, devPlan: plan }, model.revision);
      props.onSaved(m);
      setPlan(readDevPlan(m.info.devPlan));
      setMsg("Saved.");
      window.setTimeout(() => setMsg(""), 2500);
    } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  const rows: Row[] = events.map((e) => {
    const [done, all] = progress(e);
    return { id: e.id, label: e.name, sub: dates(eventSpan(e)) + (all ? ` · ${done}/${all}` : ""), span: eventSpan(e), color: e.color };
  });
  const newTask = editTask && ev && !ev.tasks.some((t) => t.id === editTask.id);

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Development plan <span className="cnt">{events.length}</span></h2>
        <div className="pi-legend">
          {dirty && <span className="pi-dirty">Unsaved</span>}
          {msg && <span className="pi-msg">{msg}</span>}
          <button className="btn sm" disabled={!!editEvent} onClick={() => setEditEvent(blankEvent(events.length))}>+ Event</button>
          <button className="btn sm primary" disabled={!dirty || busy} onClick={save}>{busy ? "Saving…" : "Save"}</button>
        </div>
      </div>
      <div className="pi-sec-b">
        {editEvent && <EventEdit key={editEvent.id} value={editEvent} onKeep={keepEvent} onCancel={() => setEditEvent(null)} />}
        {!events.length && !editEvent && <div className="muted">No event yet. Add the events of the development (ER, DR, DV, PV builds, tooling, …) with + Event.</div>}
        {events.length > 0 && (days.length
          ? <Gantt rows={rows} days={days} selected={selected} onPick={(id) => { setSelected(id); setEditTask(null); }} />
          : <div className="muted">Give the events (or their tasks) dates to see them on the calendar.</div>)}
        {events.length > 0 && !days.length && (
          <div className="wk-dp-chips">{events.map((e) => <button key={e.id} className={"wk-chip" + (e.id === selected ? " on" : "")} onClick={() => setSelected(e.id)}>{e.name}</button>)}</div>
        )}

        {ev && (
          <div className="wk-dp-detail">
            <div className="wk-dp-dh">
              <span className="dot" style={{ background: ev.color }} />
              <h3>{ev.name}</h3>
              <span className="muted small">{dates(eventSpan(ev))} · {progress(ev)[0]}/{progress(ev)[1]} done</span>
              <span className="wk-dp-act">
                <button className="wk-link" onClick={() => move(ev, -1)} disabled={events[0] === ev} aria-label="Move up">↑</button>
                <button className="wk-link" onClick={() => move(ev, 1)} disabled={events.at(-1) === ev} aria-label="Move down">↓</button>
                <button className="wk-link" disabled={!!editEvent} onClick={() => setEditEvent(ev)}>Edit event</button>
                <button className="wk-link bad" onClick={() => removeEvent(ev)}>Delete event</button>
              </span>
            </div>
            <button className="wk-qty-h" onClick={() => setShowQty(!showQty)} aria-expanded={showQty}>
              <span className="tw">{showQty ? "▾" : "▸"}</span> Quantity plan
              <span className="muted small"> · {ev.rows.length} row(s){ev.rows.length ? ` · total ${num(ev.rows.reduce((a, r) => a + rowTotal(r), 0))}` : ""}</span>
            </button>
            {showQty && <QtyTable key={ev.id} ev={ev} onRows={setRows} />}
            <div className="wk-qty-h static">Detailed schedule</div>
            {evDays.length > 0 && ev.tasks.length > 0 && (
              <Gantt days={evDays} rows={ev.tasks.map((t) => ({ id: t.id, label: t.name, sub: dates(taskSpan(t)) + (t.owner ? " · " + t.owner : ""), span: taskSpan(t), color: ev.color, done: t.done }))} />
            )}
            <table className="wk-it wk-dp-tasks">
              <thead>
                <tr><th className="fit">Done</th><th className="l">Task</th><th className="fit">Start</th><th className="fit">End</th><th className="fit">Owner</th><th className="l">Note</th><th className="fit"></th></tr>
              </thead>
              <tbody>
                {ev.tasks.map((t) => (editTask?.id === t.id
                  ? <TaskEdit key={t.id} value={editTask} onKeep={keepTask} onCancel={() => setEditTask(null)} />
                  : (
                    <tr key={t.id} className={t.done ? "done" : ""}>
                      <td><input type="checkbox" checked={t.done} aria-label={`${t.name} done`} onChange={(x) => setTasks(ev.tasks.map((y) => (y.id === t.id ? { ...y, done: x.target.checked } : y)))} /></td>
                      <td className="l"><b>{t.name}</b></td>
                      <td>{t.start || "—"}</td>
                      <td>{t.end || "—"}</td>
                      <td>{t.owner || "—"}</td>
                      <td className="l muted">{t.note}</td>
                      <td className="act">
                        <button className="wk-link" disabled={!!editTask} onClick={() => setEditTask(t)}>Edit</button>
                        <button className="wk-link bad" disabled={!!editTask} onClick={() => setTasks(ev.tasks.filter((y) => y.id !== t.id))}>Delete</button>
                      </td>
                    </tr>
                  )))}
                {newTask && <TaskEdit key={editTask!.id} value={editTask!} onKeep={keepTask} onCancel={() => setEditTask(null)} />}
                {!ev.tasks.length && !newTask && <tr><td colSpan={7} className="wk-it-empty">No detailed schedule yet.</td></tr>}
              </tbody>
            </table>
            {!editTask && <button className="btn sm wk-dp-add" onClick={() => setEditTask(blankTask())}>+ Task</button>}
          </div>
        )}
      </div>
    </section>
  );
}
