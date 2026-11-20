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
  groupQty, isMaterialTask, markDone, planDays, planProblem, progress, readDevPlan, readQty, removeTask, rowTotal, setQty, setTaskQty,
  sampleKind, sampleSubs, taskSpan, type TaskKind,
} from "./plan";
import type { DevPlan as Plan, PlanEvent, PlanTask, QtyGroup, QtyRow, TaskMaterial } from "./plan";
import { applyTidy, freeEvent, freeTask, itemOf, templateTasks, tidyRows, withCode, withStage, type TidyRow } from "./planNames";
import { builtinTask, DEFAULT_VOCAB, planVocabApi, type PlanVocab } from "./planVocab";
import MaterialPick from "./MaterialPick";
import { readMaterials, type MaterialList } from "./materialList";
import type { WorkingModel } from "./types";

interface Row {
  id: string; label: string; sub: string; span: [string, string] | null; color: string; done?: boolean; child?: boolean;
  /**
   * A row counted by day (a sample task): its day cells are the counts, typed in place (onQty), or only shown (a task
   * whose sub-tasks are counted: their sum); its days underlined in its color.
   */
  qty?: Record<string, number>;
  onQty?: (day: string, n: number | undefined) => void;
}

export const dates = (span: [string, string] | null) => (span ? (span[0] === span[1] ? span[0].slice(5) : `${span[0].slice(5)} – ${span[1].slice(5)}`) : "no date");

/** A hex color (#rgb or #rrggbb) mixed with white by `k` (0 = as is, 1 = white); any other color as is. */
export function tint(color: string, k: number): string {
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(color);
  if (!m) return color;
  const hex = m[1].length === 3 ? [...m[1]].map((c) => c + c).join("") : m[1];
  return "#" + [0, 2, 4].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) * (1 - k) + 255 * k).toString(16).padStart(2, "0")).join("");
}

/**
 * Horizontal calendar: a column per day, a row per item, the days of each item filled in its color. `expand` (rows of
 * the same day columns) opens right under the selected row; its day columns are then wide enough for a quantity
 * ("1,000"), so the calendar scrolls instead of fitting the width. Rows counted by day keep the calendar fitting the
 * width (counts set compactly in their day cells), so it scrolls only when it is long.
 */
export function Gantt(props: { rows: Row[]; days: string[]; selected?: string; onPick?: (id: string) => void; expand?: ReactNode }) {
  const tr = useTr();
  const today = localToday();
  const { rows, days } = props;
  const wide = !!props.expand;
  const counts = rows.some((r) => r.qty);
  return (
    <div className="wk-gantt-wrap">
      {/* a short calendar fills the width; a long one, or one with quantities open, scrolls */}
      <table className={"wk-gantt" + (wide ? " qopen" : days.length <= 60 ? " fill" : "") + (counts ? " counts" : "")}>
        <thead>
          <tr>
            <th className="nm" rowSpan={2}></th>
            {/* a month of a day or two: only its month, so it does not run into the next */}
            {monthRuns(days).map(([m, n]) => <th key={m} colSpan={n} className="mon" title={m}>{n < 3 ? m.slice(5) : m}</th>)}
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
          {rows.map((r) => {
            // a sub-task: a smaller row with a thinner bar in a lighter shade of its task's color
            const color = r.child ? tint(r.done ? "#94a3b8" : r.color, 0.45) : r.done ? "#94a3b8" : r.color;
            return (
            <Fragment key={r.id}>
              <tr className={(props.onPick ? "pick" : "") + (r.child ? " child" : "") + (r.id === props.selected ? " on" : "") + (r.id === props.selected && props.expand ? " open" : "")}
                onClick={props.onPick ? () => props.onPick!(r.id) : undefined}>
                <td className={"nm" + (r.child ? " child" : "")}>
                  {props.onPick && <span className="tw">{r.id === props.selected && props.expand ? "▾" : "▸"}</span>}
                  <span className="dot" style={{ background: r.child ? color : r.color }} />
                  <b className={r.done ? "done" : ""}>{tr(r.label)}</b>
                  <small>{r.sub}</small>
                </td>
                {days.map((d) => {
                  const on = inSpan(d, r.span);
                  const cls = "c" + (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "")
                    + (on ? " on" + (d === r.span![0] ? " s" : "") + (d === r.span![1] ? " e" : "") : "");
                  if (r.qty) {
                    return (
                      <td key={d} className={"q" + (r.onQty ? "" : " sum") + (d === today ? " today" : weekday(d) % 6 === 0 ? " we" : "") + (on ? " in" : "") + (r.qty[d] ? " has" : "")}
                        style={on ? { boxShadow: `inset 0 -3px 0 ${color}` } : undefined}>
                        {r.onQty
                          ? <QtyCell value={r.qty[d]} label={`${r.label} ${d}`} onKeep={(n) => r.onQty!(d, n)} onPaste={() => false} />
                          : r.qty[d] ? num(r.qty[d]) : null}
                      </td>
                    );
                  }
                  if (r.child) return <td key={d} className={cls}>{on && <i className="bar" style={{ background: color }} />}</td>;
                  return <td key={d} className={cls} style={on ? { background: color } : undefined} />;
                })}
              </tr>
              {r.id === props.selected && props.expand}
            </Fragment>
            );
          })}
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

/** Stage (a standard one with its quantity, or a name of its own), dates and color of an event being added or changed. */
export function EventEdit(props: { value: PlanEvent; onKeep: (e: PlanEvent) => void; onCancel: () => void; vocab?: PlanVocab }) {
  const vocab = props.vocab ?? DEFAULT_VOCAB;
  const [e, setE] = useState(props.value);
  const stage = itemOf(e.stage, vocab.stages);
  const kept = stage ? withStage(e, stage, (e.qty ?? "").trim()) : { ...e, name: e.name.trim() };
  const problem = planProblem(kept);
  return (
    <div className="wk-dp-form">
      <select className="inp sm" value={stage?.code ?? ""} aria-label="Stage"
        onChange={(x) => { const s = itemOf(x.target.value, vocab.stages); setE(s ? withStage(e, s, e.qty ?? "") : freeEvent(e)); }}>
        {vocab.stages.map((s) => <option key={s.code} value={s.code}>{s.name}</option>)}
        <option value="">Other (own name)</option>
      </select>
      {stage
        ? <input className="inp sm wk-dp-qty" value={e.qty ?? ""} placeholder="Qty, e.g. 1.2k" aria-label="Quantity" onChange={(x) => setE(withStage(e, stage, x.target.value))} />
        : <input className="inp sm" value={e.name} placeholder="Event, e.g. DR build" autoFocus aria-label="Event name" onChange={(x) => setE({ ...e, name: x.target.value })} />}
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
      {problem && kept.name && <span className="pi-dirty small">{problem}</span>}
      <button className="btn sm" onClick={props.onCancel}>Cancel</button>
      <button className="btn sm primary" disabled={!!problem} onClick={() => props.onKeep(kept)}>Keep</button>
    </div>
  );
}

/** A task as it is kept from its edit row (the standard name with its detail, or its own name; owner and note trimmed), and what stops keeping it. */
export function keptTask(t: PlanTask, vocab: PlanVocab, tasks: PlanTask[] = []): { task: PlanTask; problem: string } {
  const kind: TaskKind = (c) => builtinTask(c, vocab.tasks);
  const std = itemOf(t.code, vocab.tasks);
  const kept = std ? withCode(t, std, (t.detail ?? "").trim()) : { ...t, name: t.name.trim() };
  const problem = planProblem(kept) || (t.actualStart && t.actualEnd && t.actualEnd < t.actualStart ? "Actual end is before actual start." : "");
  // materials stay with a material task only, the counts of the days with a sample task (or its sub-task) only
  const { materials, qty, ...rest } = kept;
  return {
    task: {
      ...rest, ...(isMaterialTask(kept, kind) && materials?.length ? { materials } : {}), ...(sampleKind(kept, tasks, kind) && qty && Object.keys(qty).length ? { qty } : {}),
      owner: t.owner.trim(), note: t.note.trim(),
    },
    problem,
  };
}

/** The task with these materials picked (none: the field goes). */
export const withMaterials = (t: PlanTask, list: TaskMaterial[]): PlanTask => {
  const { materials: _, ...rest } = t;
  return list.length ? { ...rest, materials: list } : rest;
};

/** The name a picked material is shown by, and its sub-task named. */
export const materialLabel = (m: TaskMaterial) => m.code || m.name;

/** The task the material rows of a material task go under: the task itself, or for a sub-task (one level deep) its parent. */
const materialHost = (t: PlanTask) => t.parentId ?? t.id;

/**
 * The open edit rows with the sub-tasks of a material task following the materials picked for it (before: as it was,
 * after: with the new pick): a new sub-task row for each material that has none among the kept tasks and open rows yet,
 * named as the material, so its own dates are typed under the task (under its parent for a sub-task, beside it); a new
 * row of a material no longer picked taken out while nothing is typed in it (a kept sub-task stays to be deleted by hand).
 */
export function materialRows(open: PlanTask[], before: PlanTask, after: PlanTask, kept: PlanTask[]): PlanTask[] {
  const host = materialHost(after);
  const picked = [...new Set((after.materials ?? []).map(materialLabel))];
  const dropped = (before.materials ?? []).map(materialLabel).filter((n) => !picked.includes(n));
  const rows = open.filter((s) => !(s.id !== after.id && s.parentId === host && dropped.includes(s.name.trim()) && untouched(s) && !kept.some((k) => k.id === s.id)));
  const have = [...kept, ...rows].filter((s) => s.id !== after.id && s.parentId === host).map((s) => s.name.trim());
  return [...rows, ...picked.filter((n) => !have.includes(n)).map((name) => ({ ...blankTask(), name, parentId: host }))];
}

/** A row with nothing typed in it but its name. */
const untouched = (s: PlanTask) => !s.start && !s.end && !s.actualStart && !s.actualEnd && !s.done && !s.code && !s.owner.trim() && !s.note.trim();

/** The open edit rows with tasks opened by Edit: a material task with a new row for each picked material that has no sub-task yet. */
export const editRows = (open: PlanTask[], more: PlanTask[], kept: PlanTask[]) =>
  more.reduce((list, t) => (t.materials?.length ? materialRows(list, t, t, kept) : list), openRows(open, more));

/**
 * The open edit rows without a cancelled one: a task not kept yet goes with its new sub-task rows, a kept one (or a
 * sub-task) with the new rows of its materials still untouched.
 */
export function cancelRows(open: PlanTask[], t: PlanTask, kept: PlanTask[]): PlanTask[] {
  const isNew = (s: PlanTask) => !kept.some((k) => k.id === s.id);
  const mats = (t.materials ?? []).map(materialLabel);
  return open.filter((s) => s.id !== t.id && !(s.parentId === materialHost(t) && isNew(s) && ((isNew(t) && !t.parentId) || (mats.includes(s.name.trim()) && untouched(s)))));
}

/** The materials picked for a task, shown under its name; a button to pick them (onPick) on a material task. */
export function TaskMaterials(props: { task: PlanTask; onPick?: () => void }) {
  const tr = useTr();
  const list = props.task.materials ?? [];
  if (!props.onPick && !list.length) return null;
  return (
    <span className="wk-dp-mats">
      {list.map((m) => <span key={(m.code || m.name)} className="wk-dp-mat" title={m.name}>{m.code || m.name}</span>)}
      {props.onPick && (
        <button type="button" className="wk-link" title="Pick the materials of this task from Model Info › Material" onClick={props.onPick}>
          {list.length ? tr("Change materials") : "+ " + tr("Materials")}
        </button>
      )}
    </span>
  );
}

/** The tasks one Edit opens together: a task with its sub-tasks, a sub-task alone. */
export const editGroup = (tasks: PlanTask[], id: string) => tasks.filter((t) => t.id === id || t.parentId === id);

/** The open edit rows with more opened: a row already open stays as it is being typed. */
export const openRows = (open: PlanTask[], more: PlanTask[]) => [...open, ...more.filter((t) => !open.some((o) => o.id === t.id))];

/**
 * The schedule with rows put in (open edit rows to show, or kept ones to save): a task already there in its place, a new
 * task at the end, a new sub-task after the last sub-task of its parent.
 */
export function withRows(tasks: PlanTask[], rows: PlanTask[]): PlanTask[] {
  let out = tasks.map((t) => rows.find((r) => r.id === t.id) ?? t);
  for (const r of rows) if (!tasks.some((t) => t.id === r.id)) out = r.parentId ? addSubTask(out, r.parentId, r) : [...out, r];
  return out;
}

/**
 * One task of the detailed schedule being added or changed, a row of the table with its own Cancel and Keep; what is
 * typed is kept by the caller, so several rows (tasks, sub-tasks, new ones) can be open at once.
 */
export function TaskEditRow(props: {
  value: PlanTask; onChange: (t: PlanTask) => void; onKeep: (t: PlanTask) => void; onCancel: () => void; vocab: PlanVocab;
  /** The model's material list (Model Info › Material) a material task picks from. */
  materials?: MaterialList;
  /** The tasks of the event (a sub-task of a sample task keeps its counts). */
  tasks?: PlanTask[];
  /** The task with its materials picked (the caller adds their sub-task rows); else it is only changed. */
  onMaterials?: (t: PlanTask) => void;
  /** Why it cannot be kept yet besides its own problem (a sub-task of a task not kept yet that has one). */
  blocked?: string;
}) {
  const t = props.value;
  const kind: TaskKind = (c) => builtinTask(c, props.vocab.tasks);
  // choosing the material category opens the popup to pick its materials
  const [picking, setPicking] = useState(false);
  const setT = props.onChange;
  const set = (f: Partial<PlanTask>) => setT({ ...t, ...f });
  const std = itemOf(t.code, props.vocab.tasks);
  const kept = keptTask(t, props.vocab, props.tasks);
  const { task } = kept, problem = kept.problem || props.blocked || "";
  return (
    <tr className={"wk-dp-edit" + (t.parentId ? " sub" : "")}>
      <td></td>
      <td><input type="checkbox" checked={t.done} aria-label="Done" onChange={(x) => setT(markDone(t, x.target.checked, localToday()))} /></td>
      <td className="l wk-dp-cat">
        <select className="inp sm" value={std?.code ?? ""} aria-label="Category"
          onChange={(x) => {
            const s = itemOf(x.target.value, props.vocab.tasks);
            const n = s ? withCode(t, s, t.detail ?? "") : freeTask(t);
            setT(n);
            if (isMaterialTask(n, kind) && !isMaterialTask(t, kind)) setPicking(true);
          }}>
          <option value="">Other (own name)</option>
          {props.vocab.tasks.map((s) => <option key={s.code} value={s.code}>{s.name}</option>)}
        </select>
      </td>
      <td className="l">
        {t.parentId && <span className="wk-dp-sub">↳</span>}
        <span className="wk-dp-name">
          {std
            ? <input className="inp sm" value={t.detail ?? ""} placeholder="Detail, e.g. 300ea Bringup" aria-label="Detail" onChange={(x) => setT(withCode(t, std, x.target.value))} />
            : <input className="inp sm" value={t.name} placeholder={t.parentId ? "Sub-task" : "Task"} autoFocus={!t.name} aria-label="Task name" onChange={(x) => set({ name: x.target.value })} />}
        </span>
        {isMaterialTask(t, kind) && <TaskMaterials task={t} onPick={() => setPicking(true)} />}
        {picking && (
          <MaterialPick title={t.name || "Material in"} materials={props.materials ?? {}} value={t.materials ?? []}
            onPick={(list) => { (props.onMaterials ?? setT)(withMaterials(t, list)); setPicking(false); }} onClose={() => setPicking(false)} />
        )}
      </td>
      <td><DateBox className="sm" value={t.start} label="Start" onChange={(start) => set({ start })} /></td>
      <td><DateBox className="sm" value={t.end} label="End" onChange={(end) => set({ end })} /></td>
      <td><DateBox className="sm" value={t.actualStart ?? ""} label="Actual start" onChange={(actualStart) => set({ actualStart })} /></td>
      <td><DateBox className="sm" value={t.actualEnd ?? ""} label="Actual end" onChange={(actualEnd) => set({ actualEnd })} /></td>
      <td><input className="inp sm wk-dp-owner" value={t.owner} placeholder="Owner" onChange={(x) => set({ owner: x.target.value })} /></td>
      <td className="l"><input className="inp sm" value={t.note} placeholder="Note" onChange={(x) => set({ note: x.target.value })} /></td>
      <td className="act">
        {problem && task.name && <div className="pi-dirty small">{problem}</div>}
        <button className="btn sm" onClick={props.onCancel}>Cancel</button>{" "}
        <button className="btn sm primary" disabled={!!problem} onClick={() => props.onKeep(task)}>Keep</button>
      </td>
    </tr>
  );
}

/** One task edited on its own, keeping what is typed itself. */
export function TaskEdit(props: { value: PlanTask; onKeep: (t: PlanTask) => void; onCancel: () => void; vocab: PlanVocab; materials?: MaterialList }) {
  const [t, setT] = useState(props.value);
  return <TaskEditRow value={t} onChange={setT} onKeep={props.onKeep} onCancel={props.onCancel} vocab={props.vocab} materials={props.materials} />;
}

/** Waiting this long after the last change before saving it, so typing in the quantities is saved once. */
export const AUTO_SAVE_MS = 800;

/** Whether the plan as it is now is to be saved by itself: changed, nothing saving, and not the plan that just failed to save. */
export const autoSaveDue = (plan: Plan, saved: Plan, busy: boolean, failed: string) => {
  const now = JSON.stringify(plan);
  return !busy && now !== JSON.stringify(saved) && now !== failed;
};

/** The plan after a save: the one the server kept, unless it was changed again while saving (that change is saved next). */
export const afterSave = (current: Plan, sent: Plan, kept: Plan) => (current === sent ? kept : current);

/**
 * Development plan of the model: the events on a horizontal calendar (a row per event, its days filled in), and under
 * it the detailed schedule of the event clicked. Kept in the model information (info.devPlan) and saved here.
 */
export default function DevPlan(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
  /** Standard names (tests); else asked of the server. */
  vocab?: PlanVocab;
}) {
  const tr = useTr();
  const { model } = props;
  const [vocab, setVocab] = useState<PlanVocab>(props.vocab ?? DEFAULT_VOCAB);
  useEffect(() => { if (!props.vocab) planVocabApi.get().then(setVocab, () => { /* the default list stays */ }); }, []);
  const kind: TaskKind = (c) => builtinTask(c, vocab.tasks);
  const [tidy, setTidy] = useState<TidyRow[] | null>(null);
  const savedPlan = useMemo(() => readDevPlan(model.info.devPlan), [model]);
  const [plan, setPlan] = useState<Plan>(savedPlan);
  const [selected, setSelected] = useState(savedPlan.events[0]?.id ?? "");
  const [editEvent, setEditEvent] = useState<PlanEvent | null>(null);
  /** The edit rows open in each event (tasks being changed and new ones), as typed so far; any number at once. */
  const [open, setOpenAll] = useState<Record<string, PlanTask[]>>({});
  const [showQty, setShowQty] = useState(false);
  const [dragTask, setDragTask] = useState("");
  /** The task whose materials are being picked from its row. */
  const [pickFor, setPickFor] = useState("");
  const materials = useMemo(() => readMaterials(model.info.materials), [model]);
  const [dropOn, setDropOn] = useState("");
  const endDrag = () => { setDragTask(""); setDropOn(""); };
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  /** The plan (as JSON) whose save failed: not tried again by itself until it is changed. */
  const [failed, setFailed] = useState("");
  const dirty = JSON.stringify(plan) !== JSON.stringify(savedPlan);
  const events = plan.events;
  const ev = events.find((e) => e.id === selected) ?? null;
  const rowsOpen = ev ? open[ev.id] ?? [] : [];
  const setOpen = (f: (list: PlanTask[]) => PlanTask[]) => {
    if (!ev) return;
    const id = ev.id;
    setOpenAll((o) => ({ ...o, [id]: f(o[id] ?? []) }));
  };
  // rows not kept yet are not saved, so leaving the model with one open asks first
  const notKept = events.reduce((n, e) => n + (open[e.id]?.length ?? 0), 0);
  useEffect(() => props.onDirty(dirty || notKept > 0), [dirty, notKept, model]);
  const days = useMemo(() => planDays(events), [events]);
  const evDays = useMemo(() => (ev
    ? calendarDays([eventSpan(ev), ...ev.tasks.map(taskSpan), ...ev.tasks.flatMap((t) => Object.keys(t.qty ?? {}).map((d): [string, string] => [d, d]))], 1)
    : []), [ev]);

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
  const keptOpen = rowsOpen.map((t) => keptTask(t, vocab, ev?.tasks));
  // a sub-task row of a task not kept yet (a new material task's material) is kept with its task
  const newParent = (t: PlanTask) => (t.parentId && !ev?.tasks.some((y) => y.id === t.parentId) ? keptOpen.find((k) => k.task.id === t.parentId) : undefined);
  const keepTask = (t: PlanTask) => {
    if (!ev) return;
    const p = newParent(t);
    if (p?.problem) return;
    const rows = p ? [p.task, t] : [t];
    setTasks(withRows(ev.tasks, rows));
    setOpen((list) => list.filter((x) => !rows.some((r) => r.id === x.id)));
  };
  // the materials picked for a task bring a sub-task row for each under it
  const pickMaterials = (before: PlanTask, after: PlanTask) => ev && setOpen((list) => materialRows(list.map((x) => (x.id === after.id ? after : x)), before, after, ev.tasks));
  const keepAll = () => {
    if (!ev || keptOpen.some((k) => k.problem)) return;
    setTasks(withRows(ev.tasks, keptOpen.map((k) => k.task)));
    setOpen(() => []);
  };
  const save = async () => {
    const sent = plan;
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, { ...model.info, devPlan: sent }, model.revision);
      props.onSaved(m);
      setPlan((cur) => afterSave(cur, sent, readDevPlan(m.info.devPlan)));
      setFailed("");
      setMsg("Saved.");
      window.setTimeout(() => setMsg(""), 2500);
    } catch (e) { setFailed(JSON.stringify(sent)); props.onError(e); } finally { setBusy(false); }
  };
  // every change is saved by itself a moment later (a failed save waits for Save or the next change)
  useEffect(() => {
    if (!autoSaveDue(plan, savedPlan, busy, failed)) return;
    const t = window.setTimeout(() => void save(), AUTO_SAVE_MS);
    return () => window.clearTimeout(t);
  }, [plan, savedPlan, busy, failed]);

  const freeNames = events.reduce((n, e) => n + (e.stage ? 0 : 1) + e.tasks.filter((t) => !t.code).length, 0);
  const standard = ev ? templateTasks(ev, vocab) : [];
  const rows: Row[] = events.map((e) => {
    const [done, all] = progress(e);
    return { id: e.id, label: e.name, sub: dates(eventSpan(e)) + (all ? ` · ${done}/${all}` : ""), span: eventSpan(e), color: e.color };
  });
  // a new row is typed in its place: a sub-task right after the last sub-task of its parent, a task at the end
  const shownTasks = ev ? withRows(ev.tasks, rowsOpen) : [];
  // the samples of a sample task, with the ones of its sample sub-tasks
  const sampleTotal = (t: PlanTask) => (ev ? Object.values(groupQty(t, ev.tasks, kind)).reduce((a, b) => a + b, 0) : 0);
  const dragged = ev?.tasks.find((t) => t.id === dragTask);
  const removeOne = (t: PlanTask) => {
    const subs = ev!.tasks.filter((y) => y.parentId === t.id).length;
    if (subs && !window.confirm(`Delete the task ${t.name} and its ${subs} sub-task(s)?`)) return;
    setTasks(removeTask(ev!.tasks, t.id));
    setOpen((list) => removeTask(list, t.id));
  };
  // clicking an event opens its quantity plan right under it; clicking it again closes it
  const pick = (id: string) => {
    setShowQty(id === selected ? !showQty : true);
    setSelected(id);
  };
  const qty = ev && showQty && <QtyRows key={ev.id} ev={ev} days={days} product={model.name} onRows={setRows} onBom={(b, p) => fillBom(ev.id, b, p)} />;

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr("Development plan")} <span className="cnt">{events.length}</span></h2>
        <div className="pi-legend">
          {dirty && (failed === JSON.stringify(plan) ? <span className="pi-dirty">{tr("Not saved")}</span> : <span className="pi-msg">{tr("Saving…")}</span>)}
          {msg && !dirty && <span className="pi-msg">{msg}</span>}
          {notKept > 0 && <span className="pi-dirty" title="Rows still open are saved once kept">{notKept} {tr("row(s) not kept")}</span>}
          {freeNames > 0 && (
            <button className="btn sm" disabled={!!tidy} title="Give the events and tasks written in their own words the standard names"
              onClick={() => setTidy(tidyRows(plan, vocab))}>{tr("Standard names")} <span className="cnt">{freeNames}</span></button>
          )}
          <button className="btn sm" disabled={!!editEvent} onClick={() => setEditEvent(blankEvent(events.length))}>+ Event</button>
          <button className="btn sm primary" disabled={!dirty || busy} onClick={save}
            title="Changes are saved by themselves; Save tries again after a failed save">{busy ? "Saving…" : "Save"}</button>
        </div>
      </div>
      <div className="pi-sec-b">
        {tidy && <TidyNames rows={tidy} vocab={vocab} onRows={setTidy} onApply={() => { setPlan(applyTidy(plan, tidy, vocab)); setTidy(null); }} onCancel={() => setTidy(null)} />}
        {editEvent && <EventEdit key={editEvent.id} value={editEvent} onKeep={keepEvent} onCancel={() => setEditEvent(null)} vocab={vocab} />}
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
              <Gantt days={evDays}
                rows={ev.tasks.map((t) => {
                  const total = sampleTotal(t);
                  return {
                    id: t.id, label: t.name, sub: dates(groupSpan(ev.tasks, t)) + (t.owner ? " · " + t.owner : "") + (total ? ` · Σ ${num(total)}` : ""),
                    span: groupSpan(ev.tasks, t), color: ev.color, done: t.done, child: !!t.parentId,
                    // a sample task (or its sub-task): how many each day, typed in its day cells; one whose sub-tasks are
                    // counted shows their sum
                    ...(!sampleKind(t, ev.tasks, kind) ? {}
                      : sampleSubs(t, ev.tasks, kind).length ? { qty: groupQty(t, ev.tasks, kind) }
                      : { qty: t.qty ?? {}, onQty: (d: string, n: number | undefined) => setTasks(ev.tasks.map((y) => (y.id === t.id ? setTaskQty(y, d, n) : y))) }),
                  };
                })} />
            )}
            <table className="wk-it wk-dp-tasks">
              <thead>
                <tr><th className="fit"></th>{(["Done", "Category", "Task", "Start", "End", "Actual start", "Actual end", "Owner", "Note"] as const).map((h) => <th key={h} className={h === "Task" || h === "Note" ? "l" : "fit"}>{tr(h)}</th>)}<th className="fit"></th></tr>
              </thead>
              <tbody>
                {shownTasks.map((t) => {
                  const o = rowsOpen.find((x) => x.id === t.id);
                  return o
                    ? <TaskEditRow key={t.id} value={o} vocab={vocab} materials={materials} tasks={shownTasks} onKeep={keepTask}
                      onMaterials={(n) => pickMaterials(o, n)} blocked={newParent(o)?.problem ? "Its task above is not right yet." : ""}
                      onChange={(n) => setOpen((list) => list.map((x) => (x.id === n.id ? n : x)))}
                      onCancel={() => setOpen((list) => cancelRows(list, o, ev.tasks))} />
                    : (
                    <tr key={t.id} className={(t.parentId ? "sub" : "") + (t.done ? " done" : "") + (dropOn === t.id && dragTask ? (ev.tasks.findIndex((y) => y.id === dragTask) < ev.tasks.indexOf(t) ? " drop-below" : " drop-above") : "")}
                      draggable onDragStart={(x) => { x.dataTransfer.effectAllowed = "move"; x.dataTransfer.setData("text/plain", t.name); setDragTask(t.id); }}
                      onDragOver={(x) => {
                        // a task goes among the tasks, a sub-task among the sub-tasks of its parent
                        if (!dragged || (dragged.parentId ?? "") !== (t.parentId ?? "")) return;
                        x.preventDefault();
                        setDropOn(t.id);
                      }}
                      onDrop={(x) => { x.preventDefault(); if (dragTask) setTasks(moveTask(ev.tasks, dragTask, t.id)); endDrag(); }}
                      onDragEnd={endDrag}>
                      <td className="wk-dp-grip" title="Drag to move">⠿</td>
                      <td><input type="checkbox" checked={t.done} aria-label={`${t.name} done`} onChange={(x) => setTasks(ev.tasks.map((y) => (y.id === t.id ? markDone(y, x.target.checked, localToday()) : y)))} /></td>
                      <td className="l wk-dp-cat">{itemOf(t.code, vocab.tasks) ? <span className={"wk-dp-catv" + (sampleKind(t, ev.tasks, kind) ? " sample" : "")}>{tr(itemOf(t.code, vocab.tasks)!.name)}</span> : <span className="muted">—</span>}</td>
                      <td className="l">
                        {t.parentId && <span className="wk-dp-sub">↳</span>}<b>{tr(t.name)}</b>
                        <TaskMaterials task={t} onPick={isMaterialTask(t, kind) ? () => setPickFor(t.id) : undefined} />
                        {sampleKind(t, ev.tasks, kind) && <span className="wk-dp-qtot" title="Samples of the days typed in the calendar above">Σ {num(sampleTotal(t))}</span>}
                      </td>
                      <td>{t.start || "—"}</td>
                      <td>{t.end || "—"}</td>
                      <td className="wk-dp-real">{t.actualStart || "—"}</td>
                      <td className={"wk-dp-real" + (t.actualEnd && t.end && t.actualEnd > t.end ? " late" : "")}>{t.actualEnd || "—"}</td>
                      <td>{t.owner || "—"}</td>
                      <td className="l muted">{tr(t.note)}</td>
                      <td className="act">
                        {!t.parentId && <button className="wk-link" onClick={() => setOpen((list) => [...list, { ...blankTask(), parentId: t.id }])}>+ Sub</button>}
                        <button className="wk-link" title={ev.tasks.some((y) => y.parentId === t.id) ? "Edit this task and its sub-tasks together" : undefined}
                          onClick={() => setOpen((list) => editRows(list, editGroup(ev.tasks, t.id), ev.tasks))}>Edit</button>
                        <button className="wk-link bad" onClick={() => removeOne(t)}>Delete</button>
                      </td>
                    </tr>
                  );
                })}
                {!shownTasks.length && <tr><td colSpan={11} className="wk-it-empty">No detailed schedule yet.</td></tr>}
              </tbody>
            </table>
            {(() => {
              const pt = ev.tasks.find((y) => y.id === pickFor);
              return pt && (
                <MaterialPick title={pt.name} materials={materials} value={pt.materials ?? []} onClose={() => setPickFor("")}
                  onPick={(list) => {
                    const n = withMaterials(pt, list);
                    setTasks(ev.tasks.map((y) => (y.id === pt.id ? n : y)));
                    pickMaterials(pt, n);
                    setPickFor("");
                  }} />
              );
            })()}
            <div className="wk-dp-adds">
              <button className="btn sm wk-dp-add" onClick={() => setOpen((list) => [...list, blankTask()])}>+ Task</button>
              {ev.tasks.length > 1 && ev.tasks.some((y) => !rowsOpen.some((x) => x.id === y.id)) && (
                <button className="btn sm wk-dp-add" title="Change every task of this event at once" onClick={() => setOpen((list) => editRows(list, ev.tasks, ev.tasks))}>Edit all</button>
              )}
              {rowsOpen.length > 1 && (
                <>
                  <button className="btn sm wk-dp-add" onClick={() => setOpen(() => [])}>Cancel all</button>
                  <button className="btn sm primary" disabled={keptOpen.some((k) => k.problem)} title="Keep every open row (off while a row has a problem)" onClick={keepAll}>
                    Keep all ({rowsOpen.length})
                  </button>
                </>
              )}
              {standard.length > 0 && (
                <button className="btn sm wk-dp-add" title={standard.map((t) => t.name).join(", ")} onClick={() => setTasks([...ev.tasks, ...standard])}>
                  + {tr("Standard tasks of")} {itemOf(ev.stage, vocab.stages)?.name} ({standard.length})
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    </section>
  );
}

/**
 * The events and tasks written in their own words, each with the standard name it looks like (chosen or changed here) and
 * its quantity or detail; applying renames the ones given a standard name and leaves the others as they are.
 */
export function TidyNames(props: { rows: TidyRow[]; vocab: PlanVocab; onRows: (rows: TidyRow[]) => void; onApply: () => void; onCancel: () => void }) {
  const tr = useTr();
  const set = (key: string, f: Partial<TidyRow>) => props.onRows(props.rows.map((r) => (r.key === key ? { ...r, ...f } : r)));
  const chosen = props.rows.filter((r) => r.code).length;
  return (
    <div className="wk-dp-tidy">
      <div className="wk-dp-tidy-h">
        <b>{tr("Standard names")}</b>
        <span className="muted small">{tr("Each name written in its own words, with the standard name it looks like. Change what is wrong; the rest of the name stays as the detail.")}</span>
      </div>
      <table className="wk-it">
        <thead><tr><th className="fit"></th><th className="l">{tr("Written")}</th><th className="l">{tr("Standard name")}</th><th className="l">{tr("Quantity / detail")}</th></tr></thead>
        <tbody>
          {props.rows.map((r) => {
            const items = r.kind === "event" ? props.vocab.stages : props.vocab.tasks;
            return (
              <tr key={r.key} className={r.code ? "" : "muted"}>
                <td className="muted small">{r.kind === "event" ? tr("Event") : tr("Task")}</td>
                <td className="l">{r.from}</td>
                <td className="l">
                  <select className="inp sm" value={r.code} aria-label={`Standard name of ${r.from}`} onChange={(x) => set(r.key, { code: x.target.value })}>
                    <option value="">{tr("Keep as written")}</option>
                    {items.map((i) => <option key={i.code} value={i.code}>{i.name}</option>)}
                  </select>
                </td>
                <td className="l">
                  <input className="inp sm" value={r.detail} disabled={!r.code} aria-label={`Detail of ${r.from}`} onChange={(x) => set(r.key, { detail: x.target.value })} />
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <div className="wk-dp-tidy-f">
        <button className="btn sm" onClick={props.onCancel}>Cancel</button>
        <button className="btn sm primary" disabled={!chosen} onClick={props.onApply}>{tr("Apply")} {chosen}</button>
        <span className="muted small">{tr("Saved by itself once applied.")}</span>
      </div>
    </div>
  );
}
