import { Fragment, useEffect, useState } from "react";
import { productionApi } from "../productionInput/api";
import {
  MANUAL_TYPES, addMaterial, blankProcess, deleteMaterial, deleteProcess, fromProductionInput, moveProcess, processProblem, saveProcess, usage,
} from "./manual";
import type { ManualGroup as Group, ManualProcess, ManualType } from "./manual";

/** One process being added or changed: type, names, NG items (one per line) and the materials it uses per unit. */
function ProcessEdit(props: { no: number; value: ManualProcess; materials: string[]; onSave: (p: ManualProcess) => void; onCancel: () => void }) {
  const [p, setP] = useState(props.value);
  const [ng, setNg] = useState(props.value.ngItems.join("\n"));
  const draft = { ...p, ngItems: ng.split("\n") };
  const problem = processProblem(draft, props.materials);
  const unused = props.materials.filter((m) => !p.materials.some((u) => u.material.toLowerCase() === m.toLowerCase()));
  const setUse = (i: number, f: Partial<{ material: string; qtyPerUnit: number }>) =>
    setP((cur) => ({ ...cur, materials: cur.materials.map((u, j) => (j === i ? { ...u, ...f } : u)) }));
  return (
    <tr className="wk-man-edit">
      <td className="n">{props.no}</td>
      <td>
        <select className="inp sm" value={p.type} onChange={(e) => setP({ ...p, type: e.target.value as ManualType })}>
          {MANUAL_TYPES.map((t) => <option key={t}>{t}</option>)}
        </select>
      </td>
      <td className="l">
        <input className="inp sm" value={p.name} placeholder="Process name" autoFocus onChange={(e) => setP({ ...p, name: e.target.value })} />
        <input className="inp sm wk-man-local" value={p.localName} placeholder="Local name (optional)" onChange={(e) => setP({ ...p, localName: e.target.value })} />
      </td>
      <td className="l">
        <textarea className="inp sm wk-ta" rows={Math.max(3, ng.split("\n").length)} value={ng} placeholder={"One NG item per line\nGlue overflow\nMisalign"}
          onChange={(e) => setNg(e.target.value)} />
      </td>
      <td className="l">
        {p.materials.map((u, i) => (
          <div className="wk-man-use" key={i}>
            <select className="inp sm" value={u.material} onChange={(e) => setUse(i, { material: e.target.value })}>
              {[u.material, ...unused].filter(Boolean).map((m) => <option key={m}>{m}</option>)}
            </select>
            <input className="inp sm wk-man-qty" type="number" min={0} step="any" value={u.qtyPerUnit} aria-label="Qty per unit"
              onChange={(e) => setUse(i, { qtyPerUnit: Number(e.target.value) })} />
            <span className="muted small">/ unit</span>
            <button className="wk-link bad" onClick={() => setP({ ...p, materials: p.materials.filter((_, j) => j !== i) })}>✕</button>
          </div>
        ))}
        {unused.length > 0
          ? <button className="wk-link" onClick={() => setP({ ...p, materials: [...p.materials, { material: unused[0], qtyPerUnit: 1 }] })}>+ Material</button>
          : !props.materials.length && <span className="muted small">Add materials to the list first.</span>}
      </td>
      <td className="act">
        {problem && p.name.trim() && <div className="pi-dirty small">{problem}</div>}
        <button className="btn sm" onClick={props.onCancel}>Cancel</button>{" "}
        <button className="btn sm primary" disabled={!!problem} onClick={() => props.onSave(draft)}>Keep</button>
      </td>
    </tr>
  );
}

/**
 * Model group typed in Model Info for a model BMES > Make Model Group does not have: a table of the processes with
 * their NG items and the materials each uses per unit, and the material list. Kept with the model information (Save).
 */
export default function ManualGroup(props: { modelName: string; value: Group; onChange: (g: Group) => void; onError: (e: unknown) => void }) {
  const g = props.value;
  const [editing, setEditing] = useState<ManualProcess | null>(null);
  const [newMat, setNewMat] = useState("");
  const [source, setSource] = useState<{ id: string; name: string; build: string } | null>(null);

  // a Production Input model with the same name can be copied in
  useEffect(() => {
    let live = true;
    productionApi.models()
      .then((list) => { if (live) setSource(list.find((m) => m.name.trim().toLowerCase() === props.modelName.trim().toLowerCase()) ?? null); })
      .catch(() => {});
    return () => { live = false; };
  }, [props.modelName]);

  const copy = async () => {
    if (!source) return;
    if ((g.processes.length || g.materials.length)
      && !window.confirm(`Replace the processes and materials below with the ones of Production Input ${source.name} (${source.build})?`)) return;
    try { props.onChange(fromProductionInput((await productionApi.model(source.id)).setup)); setEditing(null); } catch (e) { props.onError(e); }
  };
  const addMat = () => {
    try { props.onChange(addMaterial(g, newMat)); setNewMat(""); } catch (e) { props.onError(e); }
  };
  const removeMat = (m: string) => {
    const n = usage(g, m);
    if (n && !window.confirm(`${m} is used in ${n} process${n > 1 ? "es" : ""}. Remove it from them too?`)) return;
    props.onChange(deleteMaterial(g, m));
  };
  const removeProc = (p: ManualProcess) => {
    if (window.confirm(`Delete process ${p.name}?`)) props.onChange(deleteProcess(g, p.id));
  };
  const keep = (p: ManualProcess) => { props.onChange(saveProcess(g, p)); setEditing(null); };
  const isNew = editing && !g.processes.some((p) => p.id === editing.id);

  return (
    <div className="wk-man">
      <div className="wk-man-h">
        <span className="muted small">Entered here because BMES has no model group for this model. Save the model information to keep changes.</span>
        {source && <button className="btn sm" onClick={copy}>Copy from Production Input {source.name} ({source.build})</button>}
      </div>
      <div className="wk-man-body">
        <div className="pi-wrap wk-man-procs">
          <table className="wk-it wk-man-t">
            <thead>
              <tr>
                <th className="fit">#</th><th className="fit">Type</th><th className="l">Process name</th>
                <th className="l">NG items</th><th className="l">Materials · qty per unit</th><th className="fit"></th>
              </tr>
            </thead>
            <tbody>
              {g.processes.map((p, i) => (
                <Fragment key={p.id}>
                  {editing?.id === p.id ? (
                    <ProcessEdit key={editing.id} no={i + 1} value={editing} materials={g.materials} onSave={keep} onCancel={() => setEditing(null)} />
                  ) : (
                    <tr>
                      <td className="n">{i + 1}</td>
                      <td><span className={"wk-man-type t-" + p.type.toLowerCase()}>{p.type}</span></td>
                      <td className="l"><b>{p.name}</b>{p.localName && <div className="muted small">{p.localName}</div>}</td>
                      <td className="l wk-man-list">{p.ngItems.length ? p.ngItems.map((n) => <div key={n}>{n}</div>) : <span className="muted">—</span>}</td>
                      <td className="l wk-man-list">
                        {p.materials.length
                          ? p.materials.map((u) => <div key={u.material}>{u.material} <span className="muted">× {u.qtyPerUnit}</span></div>)
                          : <span className="muted">—</span>}
                      </td>
                      <td className="act">
                        <button className="wk-link" disabled={i === 0} onClick={() => props.onChange(moveProcess(g, p.id, -1))} aria-label="Move up">↑</button>
                        <button className="wk-link" disabled={i === g.processes.length - 1} onClick={() => props.onChange(moveProcess(g, p.id, 1))} aria-label="Move down">↓</button>
                        <button className="wk-link" disabled={!!editing} onClick={() => setEditing(p)}>Edit</button>
                        <button className="wk-link bad" disabled={!!editing} onClick={() => removeProc(p)}>Delete</button>
                      </td>
                    </tr>
                  )}
                </Fragment>
              ))}
              {isNew && <ProcessEdit key={editing!.id} no={g.processes.length + 1} value={editing!} materials={g.materials} onSave={keep} onCancel={() => setEditing(null)} />}
              {!g.processes.length && !isNew && <tr><td colSpan={6} className="wk-it-empty">No process yet.</td></tr>}
            </tbody>
          </table>
          {!editing && (
            <button className="btn sm wk-man-add"
              onClick={() => setEditing(blankProcess(g.processes.length ? g.processes[g.processes.length - 1].type : "SUB"))}>+ Add process</button>
          )}
        </div>
        <div className="wk-man-mats">
          <table className="wk-it">
            <thead><tr><th className="l">Material</th><th className="fit">Used in</th><th className="fit"></th></tr></thead>
            <tbody>
              {g.materials.map((m) => (
                <tr key={m}>
                  <td className="l"><b>{m}</b></td>
                  <td className="muted">{usage(g, m)} proc</td>
                  <td className="act"><button className="wk-link bad" onClick={() => removeMat(m)} aria-label={`Remove ${m}`}>✕</button></td>
                </tr>
              ))}
              {!g.materials.length && <tr><td colSpan={3} className="wk-it-empty">No material yet.</td></tr>}
            </tbody>
          </table>
          <div className="wk-man-newmat">
            <input className="inp sm" value={newMat} placeholder="New material name" onChange={(e) => setNewMat(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") addMat(); }} />
            <button className="btn sm" disabled={!newMat.trim()} onClick={addMat}>+ Add</button>
          </div>
        </div>
      </div>
    </div>
  );
}
