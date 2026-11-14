import { useEffect, useState } from "react";
import { bmesApi as bmesTables } from "../../bmesApi";
import { addShift, addSub, allShifts, removeShift, removeSub, renameSub, subShiftCount, tempCounts } from "./tempGroup";
import type { TempGroup as Group, TempSub } from "./tempGroup";

const LIST_ID = "wk-tg-shifts";

/** Shift groups in columns, each with a remove button. */
function Shifts(props: { list: string[]; onRemove: (name: string) => void }) {
  if (!props.list.length) return null;
  return (
    <div className="wk-bg-shifts">
      {props.list.map((ls) => (
        <div key={ls} className="wk-bg-shift">
          <span className="wk-bg-sbadge">Shift Group</span><span className="nm">{ls}</span>
          <button className="wk-link bad wk-tg-x" onClick={() => props.onRemove(ls)} aria-label={`Remove ${ls}`}>✕</button>
        </div>
      ))}
    </div>
  );
}

/** A name typed once and added as a shift group (Enter) or a sub-group under this place. */
function AddRow(props: { onShift: (name: string) => boolean; onSub: (name: string) => boolean }) {
  const [name, setName] = useState("");
  const done = (ok: boolean) => { if (ok) setName(""); };
  return (
    <div className="wk-tg-add">
      <input className="inp sm" value={name} list={LIST_ID} placeholder="Shift group (e.g. MSM-A386B_LINE 1) or sub group name"
        aria-label="Shift group or sub group name" onChange={(e) => setName(e.target.value)}
        onKeyDown={(e) => { if (e.key === "Enter") done(props.onShift(name)); }} />
      <button className="btn sm" disabled={!name.trim()} onClick={() => done(props.onShift(name))}>+ Shift group</button>
      <button className="btn sm" disabled={!name.trim()} onClick={() => done(props.onSub(name))}>+ Sub group</button>
    </div>
  );
}

/** Name of a sub-group, kept when the box is left or Enter is pressed. */
function SubName(props: { value: string; onKeep: (name: string) => boolean }) {
  const [text, setText] = useState(props.value);
  useEffect(() => setText(props.value), [props.value]);
  const keep = () => { if (text.trim() !== props.value && !props.onKeep(text)) setText(props.value); };
  return (
    <input className="inp sm wk-tg-name" value={text} aria-label="Sub group name" onChange={(e) => setText(e.target.value)}
      onBlur={keep} onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); }} />
  );
}

/**
 * Temporary model group of a model BMES > Make Model Group cannot make (not in BMES yet): the Model › Sub Group ›
 * Shift Group tree as that page shows it, made here by hand and saved with the model information.
 */
export default function TempGroupEdit(props: {
  modelName: string; value: Group; onChange: (g: Group) => void; onError: (e: unknown) => void;
  /** Sub-group whose processes and materials are shown below ("" = none); clicking a sub-group picks it. */
  selected?: string; onSelect?: (id: string) => void;
  /** Processes in a sub-group's table, told before the sub-group is removed. */
  processes?: (id: string) => number;
}) {
  const g = props.value;
  const [known, setKnown] = useState<string[]>([]);

  // shift groups the NG Rate data already has for this model name, offered while typing (none when there are none)
  useEffect(() => {
    let live = true;
    bmesTables.modelGroups.lineShifts(props.modelName).then((r) => { if (live) setKnown(r.lineShifts); }).catch(() => undefined);
    return () => { live = false; };
  }, [props.modelName]);

  const apply = (f: () => Group) => {
    try { props.onChange(f()); return true; } catch (e) { props.onError(e); return false; }
  };
  const dropSub = (s: TempSub) => {
    const n = subShiftCount(s);
    const p = props.processes?.(s.id) ?? 0;
    if ((n || s.children.length || p)
      && !window.confirm(`Remove sub group ${s.name} with its ${n} shift group(s)${p ? `, ${p} process(es)` : ""}${s.children.length ? " and sub groups" : ""}?`)) return;
    apply(() => removeSub(g, s.id));
  };
  const counts = tempCounts(g);
  const used = new Set(allShifts(g).map((x) => x.toLowerCase()));

  const card = (s: TempSub, depth: number) => (
    <div key={s.id} className={"wk-bg-sub" + (props.selected === s.id ? " on" : "")} title={props.onSelect ? "Click to show its processes and materials below" : undefined}
      onClick={(e) => { e.stopPropagation(); props.onSelect?.(s.id); }}>
      <div className="wk-bg-row">
        <span className="wk-bg-badge">Sub Group {depth}</span>
        <SubName value={s.name} onKeep={(n) => apply(() => renameSub(g, s.id, n))} />
        <span className="muted small wk-bg-count">{subShiftCount(s)} shifts</span>
        <button className="wk-link bad wk-tg-x" onClick={() => dropSub(s)} aria-label={`Remove sub group ${s.name}`}>✕</button>
      </div>
      <Shifts list={s.lineShifts} onRemove={(n) => apply(() => removeShift(g, n))} />
      <AddRow onShift={(n) => apply(() => addShift(g, s.id, n))} onSub={(n) => apply(() => addSub(g, s.id, n))} />
      {s.children.length > 0 && <div className="wk-bg-subs">{s.children.map((c) => card(c, depth + 1))}</div>}
    </div>
  );

  return (
    <div className="wk-bg-model wk-tg">
      <div className="wk-bg-row">
        <span className="wk-bg-badge model">Model</span><b>{props.modelName}</b>
        <span className="wk-tg-tag">Temporary</span>
        <input className="inp sm wk-tg-grp" value={g.group} placeholder="Product group, e.g. MSM" aria-label="Product group"
          onChange={(e) => props.onChange({ ...g, group: e.target.value })} />
        <span className="muted small wk-bg-count">{counts.subs} sub groups · {counts.shifts} shifts</span>
      </div>
      <Shifts list={g.lineShifts} onRemove={(n) => apply(() => removeShift(g, n))} />
      <AddRow onShift={(n) => apply(() => addShift(g, null, n))} onSub={(n) => apply(() => addSub(g, null, n))} />
      {g.subGroups.length > 0 && <div className="wk-bg-subs">{g.subGroups.map((s) => card(s, 1))}</div>}
      <datalist id={LIST_ID}>{known.filter((x) => !used.has(x.toLowerCase())).map((x) => <option key={x} value={x} />)}</datalist>
    </div>
  );
}
