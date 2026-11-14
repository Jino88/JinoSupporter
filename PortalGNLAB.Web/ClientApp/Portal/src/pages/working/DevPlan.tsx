import { Fragment, useEffect, useMemo, useState } from "react";
import DateBox from "./DateBox";
import { useTr } from "./i18n";
import type { ReactNode } from "react";
import { WEEKDAYS, localToday, weekday } from "./dates";
import { workingApi } from "./api";
import { bmesApi } from "./bmes";
import type { BomRow } from "./bmes";
import {
  COLORS, QTY_GROUPS, addBomRows, addSubTask, blankEvent, blankRow, blankTask, calendarDays, eventSpan, groupSpan, inSpan, bomRowNames, monthRuns, moveInGroup, moveTask,
  planDays, planProblem, progress, readDevPlan, readQty, removeTask, rowTotal, setQty, taskSpan,
} from "./plan";
import type { DevPlan as Plan, PlanEvent, PlanTask, QtyGroup, QtyRow } from "./plan";
import type { WorkingModel } from "./types";

interface Row { id: string; label: string; sub: string; span: [string, string] | null; color: string; done?: boolean; child?: boolean }

export const dates = (span: [string, string] | null) => (span ? (span[0] === span[1] ? span[0].slice(5) : `${span[0].slice(5)} – ${span[1].slice(5)}`) : "no date");

/**
 * Horizontal calendar: a column per day, a row per item, the days of each item filled in its color. `expand` (rows of
 * the same day columns) opens right under the selected row; its day columns are then wide enough for a quantity
 * ("1,000"), so the calendar scrolls instead of fitting the width.
 */
export function Gantt(props: { rows: Row[]; days: string[]; selected?: string; onPick?: (id: string) => void; expand?: ReactNode }) {
  const tr = useTr();
  const today = localToday();
  const { rows, days } = props;
  return (
    <div className="wk-gantt-wrap">
      {/* a short calendar fills the width; a long one, or one with quantities open, scrolls */}
      <table className={"wk-gantt" + (props.expand ? " qopen" : days.length <= 60 ? " fill" : "")}>
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
            <Fragment key={r.id}>
              <tr className={(props.onPick ? "pick" : "") + (r.id === props.selected ? " on" : "") + (r.id === props.selected && props.expand ? " open" : "")}
                onClick={props.onPick ? () => props.onPick!(r.id) : undefined}>
                <td className={"nm" + (r.child ? " child" : "")}>
                  {props.onPick && <span className="tw">{r.id === props.selected && props.expand ? "▾" : "▸"}</span>}
                  <span className="dot" style={{ background: r.color }} />
                  <b className={r.done ? "done" : ""}>{tr(r.label)}</b>
                  <small>{r.sub}</small>
                </td>
                {days.map((d) => {
                  const on = inSpan(d, r.span);
                  const cls = "c" + (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "")
                    + (on ? " on" + (d === r.span![0] ? " s" : "") + (d === r.span![1] ? " e" : "") : "");
                  return <td key={d} className={cls} style={on ? { background: r.done ? "#94a3b8" : r.color } : undefined} />;
                })}
              </tr>
              {r.id === props.selected && props.expand}
            </Fragment>
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
 * Quantity plan of an event, as rows of the calendar under it so each day lines up with the day above, in two groups:
 * Material Plan (the materials to get in; the BOM's direct materials are put here the first time it opens) and Making
 * Plan (SPK Assy, Lower Assy, Function, OQC, ...). Each row has its total and the quantity of a day typed in its cell;
 * rows can be added, renamed, moved within their group and deleted (a material deleted is not put back unless From BOM
 * is pressed); cells copied from Excel can be pasted from the cell clicked.
 */
function QtyRows(props: {
  ev: PlanEvent; days: string[];
  /** Product whose BOM fills the Material Plan (the model name). */
  product: string;
  onRows: (rows: QtyRow[]) => void;
  onBom: (bom: BomRow[], product: string) => void;
}) {
  const tr = useTr();
  const { ev, days, product } = props;
  const today = localToday();
  const span = eventSpan(ev);
  const rows = ev.rows;
  const [busy, setBusy] = useState(false);
  const [bomMsg, setBomMsg] = useState("");

  /** The BOM of the product (local cache, else read from BMES) put into the Material Plan, ones it has left out. */
  const loadBom = async () => {
    if (!product.trim()) return;
    setBusy(true);
    setBomMsg("");
    try {
      let bom = await bmesApi.bom(product);
      if (!bom.rows.length) bom = await bmesApi.fetchBom(product);
      if (!bom.rows.length) { setBomMsg(bom.error || `No BOM found for ${product}. Add the materials with + Row.`); return; }
      const n = bomRowNames(rows, bom.rows).length;
      props.onBom(bom.rows, product);
      setBomMsg(n ? `${n} material(s) added from the BOM of ${product}.` : `Every direct material of the ${product} BOM is already here.`);
    } catch (e) {
      setBomMsg("BOM could not be read: " + (e instanceof Error ? e.message : String(e)));
    } finally { setBusy(false); }
  };
  // the first time the plan of this event opens, its materials come from the BOM
  useEffect(() => {
    if (!ev.bomFrom) void loadBom();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ev.id]);

  const setRow = (r: QtyRow) => props.onRows(rows.map((x) => (x.id === r.id ? r : x)));
  const remove = (r: QtyRow) => {
    if (Object.keys(r.qty).length && !window.confirm(`Delete the row ${r.name || "(no name)"} and its quantities?`)) return;
    props.onRows(rows.filter((x) => x.id !== r.id));
  };
  /** Cells copied from Excel (more than one) go to the right and down from the cell within its group; numbers only, blanks clear. */
  const paste = (list: QtyRow[], ri: number, di: number, text: string) => {
    const grid = text.replace(/\r?\n$/, "").split(/\r?\n/).map((l) => l.split("\t"));
    if (grid.length === 1 && grid[0].length === 1) return false;
    const changed = new Map<string, QtyRow>();
    grid.forEach((line, y) => {
      const r = list[ri + y];
      if (!r) return;
      let next = r;
      line.forEach((cell, x) => {
        const d = days[di + x];
        const n = readQty(cell);
        if (d && n !== null) next = setQty(next, d, n);
      });
      changed.set(r.id, next);
    });
    props.onRows(rows.map((r) => changed.get(r.id) ?? r));
    return true;
  };
  const dayCls = (d: string) => (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "") + (inSpan(d, span) ? " in" : "");
  const hint = (g: QtyGroup) => (g === "material"
    ? (busy ? `Reading the BOM of ${product}…` : bomMsg || "No material yet. From BOM puts the BOM's direct materials here; + Row adds one.")
    : "No item yet. Add the stages to make (SPK Assy, Lower Assy, …) with + Row.");
  return (
    <>
      {QTY_GROUPS.map(({ key: g, label }) => {
        const list = rows.filter((r) => r.group === g);
        const total = list.reduce((a, r) => a + rowTotal(r), 0);
        return (
          <Fragment key={g}>
            <tr className={"qh " + g}>
              <td className="nm">
                <div className="wk-qn">
                  <b>{tr(label)}</b>
                  <span className="muted">{list.length} row(s){total ? ` · ${num(total)}` : ""}</span>
                  {g === "material" && (
                    <button className="btn sm" disabled={busy || !product.trim()} onClick={() => void loadBom()}
                      title={`Add the direct materials (depth 1) of the ${product} BOM that are not here (local BOM cache, else read from BMES)`}>
                      {busy ? "Loading…" : "From BOM"}
                    </button>
                  )}
                  <button className="btn sm" onClick={() => props.onRows([...rows, blankRow("", g)])}>+ Row</button>
                </div>
              </td>
              {list.length
                ? days.map((d) => {
                  const sum = list.reduce((a, r) => a + (r.qty[d] ?? 0), 0);
                  return <td key={d} className={"q" + dayCls(d)} title={`${d} total`}>{sum ? num(sum) : ""}</td>;
                })
                : <td className="q hint" colSpan={days.length}>{hint(g)}</td>}
            </tr>
            {g === "material" && list.length > 0 && (busy || bomMsg) && (
              <tr className="qh"><td className="nm" /><td className="q hint" colSpan={days.length}>{hint(g)}</td></tr>
            )}
            {list.map((r, ri) => (
              <tr key={r.id} className="qr">
                <td className="nm"><div className="wk-qn">
                  <input className="wk-qty-name" value={r.name} title={r.name} placeholder={g === "material" ? "Material name" : "Item, e.g. SPK Assy"}
                    aria-label={g === "material" ? "Material name" : "Item name"} onChange={(x) => setRow({ ...r, name: x.target.value })} />
                  <span className="tot" title="Total">{rowTotal(r) ? num(rowTotal(r)) : ""}</span>
                  <span className="act">
                    <button className="wk-link" onClick={() => props.onRows(moveInGroup(rows, r.id, -1))} disabled={ri === 0} aria-label="Move up">↑</button>
                    <button className="wk-link" onClick={() => props.onRows(moveInGroup(rows, r.id, 1))} disabled={ri === list.length - 1} aria-label="Move down">↓</button>
                    <button className="wk-link bad" onClick={() => remove(r)} aria-label={`Delete ${r.name}`}>×</button>
                  </span>
                </div></td>
                {days.map((d, di) => (
                  <td key={d} className={"q" + dayCls(d) + (r.qty[d] ? " has" : "")}>
                    <QtyCell value={r.qty[d]} label={`${r.name} ${d}`} onKeep={(n) => setRow(setQty(r, d, n))} onPaste={(t) => paste(list, ri, di, t)} />
                  </td>
                ))}
              </tr>
            ))}
          </Fragment>
        );
      })}
    </>
  );
}

/** Name, dates and color of an event being added or changed. */
export function EventEdit(props: { value: PlanEvent; onKeep: (e: PlanEvent) => void; onCancel: () => void }) {
  const [e, setE] = useState(props.value);
  const problem = planProblem(e);
  return (
    <div className="wk-dp-form">
      <input className="inp sm" value={e.name} placeholder="Event, e.g. DR build" autoFocus aria-label="Event name"
        onChange={(x) => setE({ ...e, name: x.target.value })} />
      <DateBox className="sm" value={e.start} label="Start" onChange={(start) => setE({ ...e, start })} />
      <span className="muted">–</span>
      <DateBox className="sm" value={e.end} label="End" onChange={(end) => setE({ ...e, end })} />
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
    <tr className={"wk-dp-edit" + (t.parentId ? " sub" : "")}>
      <td></td>
      <td><input type="checkbox" checked={t.done} aria-label="Done" onChange={(x) => set({ done: x.target.checked })} /></td>
      <td className="l">
        {t.parentId && <span className="wk-dp-sub">↳</span>}
        <input className="inp sm" value={t.name} placeholder={t.parentId ? "Sub-task" : "Task"} autoFocus onChange={(x) => set({ name: x.target.value })} />
      </td>
      <td><DateBox className="sm" value={t.start} label="Start" onChange={(start) => set({ start })} /></td>
      <td><DateBox className="sm" value={t.end} label="End" onChange={(end) => set({ end })} /></td>
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
  const tr = useTr();
  const { model } = props;
  const savedPlan = useMemo(() => readDevPlan(model.info.devPlan), [model]);
  const [plan, setPlan] = useState<Plan>(savedPlan);
  const [selected, setSelected] = useState(savedPlan.events[0]?.id ?? "");
  const [editEvent, setEditEvent] = useState<PlanEvent | null>(null);
  const [editTask, setEditTask] = useState<PlanTask | null>(null);
  const [showQty, setShowQty] = useState(false);
  const [dragTask, setDragTask] = useState("");
  const [dropOn, setDropOn] = useState("");
  const endDrag = () => { setDragTask(""); setDropOn(""); };
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const dirty = JSON.stringify(plan) !== JSON.stringify(savedPlan);
  useEffect(() => props.onDirty(dirty), [dirty]);

  const events = plan.events;
  const ev = events.find((e) => e.id === selected) ?? null;
  const days = useMemo(() => planDays(events), [events]);
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
  // BOM materials arrive later: put them into the event as it is then, ones it has by then left out
  const fillBom = (evId: string, bom: BomRow[], product: string) =>
    setPlan((p) => ({ ...p, events: p.events.map((e) => (e.id === evId ? addBomRows(e, bomRowNames(e.rows, bom), product) : e)) }));
  const setTasks = (tasks: PlanTask[]) => ev && setEvents(events.map((x) => (x.id === ev.id ? { ...x, tasks } : x)));
  const keepTask = (t: PlanTask) => {
    if (!ev) return;
    setTasks(ev.tasks.some((x) => x.id === t.id) ? ev.tasks.map((x) => (x.id === t.id ? t : x)) : t.parentId ? addSubTask(ev.tasks, t.parentId, t) : [...ev.tasks, t]);
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
  // a new sub-task is typed in its place, right after the last sub-task of its parent
  const shownTasks = ev && newTask && editTask?.parentId ? addSubTask(ev.tasks, editTask.parentId, editTask) : ev?.tasks ?? [];
  const dragged = ev?.tasks.find((t) => t.id === dragTask);
  const removeOne = (t: PlanTask) => {
    const subs = ev!.tasks.filter((y) => y.parentId === t.id).length;
    if (subs && !window.confirm(`Delete the task ${t.name} and its ${subs} sub-task(s)?`)) return;
    setTasks(removeTask(ev!.tasks, t.id));
  };
  // clicking an event opens its quantity plan right under it; clicking it again closes it
  const pick = (id: string) => {
    setShowQty(id === selected ? !showQty : true);
    setSelected(id);
    setEditTask(null);
  };
  const qty = ev && showQty && <QtyRows key={ev.id} ev={ev} days={days} product={model.name} onRows={setRows} onBom={(b, p) => fillBom(ev.id, b, p)} />;

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr("Development plan")} <span className="cnt">{events.length}</span></h2>
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
          ? <Gantt rows={rows} days={days} selected={selected} onPick={pick} expand={qty || undefined} />
          : <div className="muted">Give the events (or their tasks) dates to see them on the calendar.</div>)}
        {events.length > 0 && !days.length && (
          <div className="wk-dp-chips">{events.map((e) => <button key={e.id} className={"wk-chip" + (e.id === selected ? " on" : "")} onClick={() => pick(e.id)}>{tr(e.name)}</button>)}</div>
        )}

        {ev && (
          <div className="wk-dp-detail">
            <div className="wk-dp-dh">
              <span className="dot" style={{ background: ev.color }} />
              <h3>{tr(ev.name)}</h3>
              <span className="muted small">{dates(eventSpan(ev))} · {progress(ev)[0]}/{progress(ev)[1]} done</span>
              <span className="wk-dp-act">
                <button className="wk-link" onClick={() => move(ev, -1)} disabled={events[0] === ev} aria-label="Move up">↑</button>
                <button className="wk-link" onClick={() => move(ev, 1)} disabled={events.at(-1) === ev} aria-label="Move down">↓</button>
                <button className="wk-link" disabled={!!editEvent} onClick={() => setEditEvent(ev)}>Edit event</button>
                <button className="wk-link bad" onClick={() => removeEvent(ev)}>Delete event</button>
              </span>
            </div>
            <div className="wk-qty-h static">{tr("Detailed schedule")}</div>
            {evDays.length > 0 && ev.tasks.length > 0 && (
              <Gantt days={evDays} rows={ev.tasks.map((t) => ({ id: t.id, label: t.name, sub: dates(groupSpan(ev.tasks, t)) + (t.owner ? " · " + t.owner : ""), span: groupSpan(ev.tasks, t), color: ev.color, done: t.done, child: !!t.parentId }))} />
            )}
            <table className="wk-it wk-dp-tasks">
              <thead>
                <tr><th className="fit"></th>{(["Done", "Task", "Start", "End", "Owner", "Note"] as const).map((h) => <th key={h} className={h === "Task" || h === "Note" ? "l" : "fit"}>{tr(h)}</th>)}<th className="fit"></th></tr>
              </thead>
              <tbody>
                {shownTasks.map((t) => (editTask?.id === t.id
                  ? <TaskEdit key={t.id} value={editTask} onKeep={keepTask} onCancel={() => setEditTask(null)} />
                  : (
                    <tr key={t.id} className={(t.parentId ? "sub" : "") + (t.done ? " done" : "") + (dropOn === t.id && dragTask ? (ev.tasks.findIndex((y) => y.id === dragTask) < ev.tasks.indexOf(t) ? " drop-below" : " drop-above") : "")}
                      draggable={!editTask} onDragStart={(x) => { x.dataTransfer.effectAllowed = "move"; x.dataTransfer.setData("text/plain", t.name); setDragTask(t.id); }}
                      onDragOver={(x) => {
                        // a task goes among the tasks, a sub-task among the sub-tasks of its parent
                        if (!dragged || (dragged.parentId ?? "") !== (t.parentId ?? "")) return;
                        x.preventDefault();
                        setDropOn(t.id);
                      }}
                      onDrop={(x) => { x.preventDefault(); if (dragTask) setTasks(moveTask(ev.tasks, dragTask, t.id)); endDrag(); }}
                      onDragEnd={endDrag}>
                      <td className="wk-dp-grip" title="Drag to move">⠿</td>
                      <td><input type="checkbox" checked={t.done} aria-label={`${t.name} done`} onChange={(x) => setTasks(ev.tasks.map((y) => (y.id === t.id ? { ...y, done: x.target.checked } : y)))} /></td>
                      <td className="l">{t.parentId && <span className="wk-dp-sub">↳</span>}<b>{tr(t.name)}</b></td>
                      <td>{t.start || "—"}</td>
                      <td>{t.end || "—"}</td>
                      <td>{t.owner || "—"}</td>
                      <td className="l muted">{tr(t.note)}</td>
                      <td className="act">
                        {!t.parentId && <button className="wk-link" disabled={!!editTask} onClick={() => setEditTask({ ...blankTask(), parentId: t.id })}>+ Sub</button>}
                        <button className="wk-link" disabled={!!editTask} onClick={() => setEditTask(t)}>Edit</button>
                        <button className="wk-link bad" disabled={!!editTask} onClick={() => removeOne(t)}>Delete</button>
                      </td>
                    </tr>
                  )))}
                {newTask && !editTask!.parentId && <TaskEdit key={editTask!.id} value={editTask!} onKeep={keepTask} onCancel={() => setEditTask(null)} />}
                {!ev.tasks.length && !newTask && <tr><td colSpan={8} className="wk-it-empty">No detailed schedule yet.</td></tr>}
              </tbody>
            </table>
            {!editTask && <button className="btn sm wk-dp-add" onClick={() => setEditTask(blankTask())}>+ Task</button>}
          </div>
        )}
      </div>
    </section>
  );
}
