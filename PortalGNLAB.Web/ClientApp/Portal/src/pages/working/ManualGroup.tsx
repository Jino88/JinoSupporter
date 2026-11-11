import { Fragment, useEffect, useState } from "react";
import { productionApi } from "../productionInput/api";
import { workingApi } from "./api";
import { bmesApi } from "./bmes";
import {
  MANUAL_TYPES, addMaterial, addMaterials, aiAnswer, bomMaterials, blankProcess, deleteMaterial, deleteProcess, fromProductionInput, moveProcess, processProblem, putMaterial, saveProcess, usage,
} from "./manual";
import type { ManualGroup as Group, ManualProcess, ManualType } from "./manual";
import { mergeSheet, parseNgHeader } from "./ngSheet";

/** The header rows of a Daily NG Rate sheet pasted here: the processes and NG items read from them, put into the table. */
function SheetPaste(props: { value: Group; onApply: (g: Group) => void; onClose: () => void }) {
  const [text, setText] = useState("");
  const sheet = text.trim() ? parseNgHeader(text) : null;
  const merged = sheet && !sheet.problem ? mergeSheet(props.value, sheet.processes) : null;
  const apply = () => {
    if (!merged) return;
    if (merged.dropped.length && !window.confirm(`These processes are not in the sheet and will be removed:\n${merged.dropped.join(", ")}`)) return;
    props.onApply(merged.group);
  };
  return (
    <div className="wk-man-paste">
      <div className="muted small">
        Copy the process row and the Input / NG / NG rate (%) row of the Daily NG Rate sheet in Excel together and paste them here.
        Processes with the same name keep their type and materials; the material list stays.
      </div>
      <textarea className="inp sm wk-ta" rows={4} value={text} autoFocus placeholder="Paste the header rows here (Ctrl+V)"
        onChange={(e) => setText(e.target.value)} />
      {sheet?.problem && <div className="pi-dirty small">{sheet.problem}</div>}
      {merged && (
        <>
          <div className="small"><b>{sheet!.processes.length} processes</b>
            {merged.dropped.length > 0 && <span className="pi-dirty"> · removes {merged.dropped.join(", ")}</span>}
          </div>
          {sheet!.warnings.map((w) => <div key={w} className="pi-dirty small">{w}</div>)}
          <div className="pi-wrap wk-man-paste-list">
            <table className="wk-it">
              <thead><tr><th className="fit">#</th><th className="l">Process name</th><th className="l">NG items</th></tr></thead>
              <tbody>
                {sheet!.processes.map((p, i) => (
                  <tr key={i}>
                    <td className="n">{i + 1}</td>
                    <td className="l"><b>{p.name}</b>{p.localName && <div className="muted small">{p.localName}</div>}</td>
                    <td className="l wk-man-list">{p.ngItems.join(" | ") || <span className="muted">—</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      <div className="wk-man-paste-b">
        <button className="btn sm" onClick={props.onClose}>Cancel</button>
        <button className="btn sm primary" disabled={!merged} onClick={apply}>Put into the table</button>
      </div>
    </div>
  );
}

/**
 * A request typed for AI (with cells copied from Excel if wanted), e.g. "fill the table from this header" or "put LOWER
 * into processes 3 to 6". The changed table it answers is shown first and goes into the table only when asked.
 */
function AiRequest(props: { modelName: string; value: Group; onApply: (g: Group) => void }) {
  const [text, setText] = useState("");
  const [busy, setBusy] = useState<AbortController | null>(null);
  const [error, setError] = useState("");
  const [answer, setAnswer] = useState<(ReturnType<typeof aiAnswer> & { note: string; model: string; asked: Group }) | null>(null);
  const send = async () => {
    if (!text.trim() || busy) return;
    const ctl = new AbortController();
    setBusy(ctl); setError(""); setAnswer(null);
    try {
      const r = await workingApi.manualAi(props.modelName, text, props.value, ctl.signal);
      setAnswer({ ...aiAnswer(props.value, r.group), note: r.note, model: r.model, asked: props.value });
    } catch (e) {
      if (!ctl.signal.aborted) setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(null); }
  };
  const apply = () => {
    if (!answer) return;
    if (answer.asked !== props.value && !window.confirm("The table changed after the request. Put the answer in anyway (those changes are replaced)?")) return;
    if (answer.removed.length && !window.confirm(`The answer removes these processes:\n${answer.removed.join(", ")}`)) return;
    props.onApply(answer.group);
    setAnswer(null); setText("");
  };
  return (
    <div className="wk-man-paste">
      <div className="muted small">
        Ask AI to change the table: type what to do and paste cells copied from Excel if needed
        (e.g. "Fill the processes and NG items from this header", "Put LOWER into processes 3 to 6"). Ctrl+Enter sends.
      </div>
      <textarea className="inp sm wk-ta" rows={3} value={text} placeholder="What should change in the table?" disabled={!!busy}
        onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void send(); } }} />
      {error && <div className="pi-dirty small">{error}</div>}
      {answer && (
        <>
          <div className="small">
            {answer.note && <div>{answer.note}</div>}
            <b>{answer.group.processes.length} processes</b> · {answer.group.materials.length} materials
            {answer.added.length > 0 && <span> · adds {answer.added.join(", ")}</span>}
            {answer.removed.length > 0 && <span className="pi-dirty"> · removes {answer.removed.join(", ")}</span>}
            <span className="muted"> · {answer.model}</span>
          </div>
          <div className="pi-wrap wk-man-paste-list">
            <table className="wk-it">
              <thead><tr><th className="fit">#</th><th className="fit">Type</th><th className="l">Process name</th><th className="l">NG items</th><th className="l">Materials</th></tr></thead>
              <tbody>
                {answer.group.processes.map((p, i) => (
                  <tr key={p.id}>
                    <td className="n">{i + 1}</td>
                    <td><span className={"wk-man-type t-" + p.type.toLowerCase()}>{p.type}</span></td>
                    <td className="l"><b>{p.name}</b>{p.localName && <div className="muted small">{p.localName}</div>}</td>
                    <td className="l wk-man-list">{p.ngItems.join(" | ") || <span className="muted">—</span>}</td>
                    <td className="l wk-man-list">{p.materials.map((u) => `${u.material} × ${u.qtyPerUnit}`).join(", ") || <span className="muted">—</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      <div className="wk-man-paste-b">
        {busy && <span className="muted small">Asking AI…</span>}
        {busy && <button className="btn sm" onClick={() => busy.abort()}>Stop</button>}
        {answer && <button className="btn sm" onClick={() => setAnswer(null)}>Discard</button>}
        {answer
          ? <button className="btn sm primary" onClick={apply}>Put into the table</button>
          : <button className="btn sm primary" disabled={!text.trim() || !!busy} onClick={() => void send()}>Ask AI</button>}
      </div>
    </div>
  );
}

/**
 * The product's direct materials (BOM level 1) put into the material list: read from the local BOM cache, or from BMES
 * when the cache has none or when asked. The new ones are listed with a check box first and added on Add.
 */
function BomImport(props: { modelName: string; value: Group; onApply: (g: Group) => void }) {
  const [product, setProduct] = useState(props.modelName);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [found, setFound] = useState<{ names: string[]; level1: number; from: string } | null>(null);
  const [off, setOff] = useState<string[]>([]);
  const load = async (fromBmes: boolean) => {
    const name = product.trim();
    if (!name || busy) return;
    setBusy(true); setError(""); setFound(null); setOff([]);
    try {
      let bom = fromBmes ? null : await bmesApi.bom(name);
      if (!bom?.rows.length) bom = await bmesApi.fetchBom(name);
      if (!bom.rows.length) { setError(bom.error || `No BOM found for ${name}.`); return; }
      setFound({
        names: bomMaterials(props.value, bom.rows),
        level1: bom.rows.filter((r) => r.level === 1).length,
        from: bom.source === "server" ? "BMES" : "the local BOM cache",
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(false); }
  };
  const take = found ? found.names.filter((n) => !off.includes(n)) : [];
  const apply = () => { props.onApply(addMaterials(props.value, take)); setFound(null); };
  return (
    <div className="wk-man-bom">
      <div className="small"><b>Load from BOM</b> <span className="muted">(direct materials, depth 1)</span></div>
      <div className="wk-man-newmat">
        <input className="inp sm" value={product} placeholder="Product name in BMES" disabled={busy}
          onChange={(e) => setProduct(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") void load(false); }} />
        <button className="btn sm" disabled={!product.trim() || busy} onClick={() => void load(false)}>{busy ? "Loading…" : "Load"}</button>
      </div>
      {error && <div className="pi-dirty small">{error}</div>}
      {found && (
        <>
          <div className="muted small">
            {found.level1} direct materials from {found.from} · {found.names.length ? `${found.names.length} not in the list` : "all already in the list"}
            {found.from !== "BMES" && <> · <button className="wk-link" onClick={() => void load(true)}>Read from BMES again</button></>}
          </div>
          {found.names.map((n) => (
            <label key={n} className="wk-man-bomi small">
              <input type="checkbox" checked={!off.includes(n)}
                onChange={(e) => setOff(e.target.checked ? off.filter((x) => x !== n) : [...off, n])} /> {n}
            </label>
          ))}
          <div className="wk-man-paste-b">
            <button className="btn sm" onClick={() => setFound(null)}>Cancel</button>
            {found.names.length > 0 && <button className="btn sm primary" disabled={!take.length} onClick={apply}>Add {take.length}</button>}
          </div>
        </>
      )}
    </div>
  );
}

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
  const [pasting, setPasting] = useState(false);
  // materials picked in the list (Ctrl+click for several), the ones dragged onto a process row, and the row they are over
  const [picked, setPicked] = useState<string[]>([]);
  const [drag, setDrag] = useState<string[] | null>(null);
  const [over, setOver] = useState<string | null>(null);

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
  const sel = g.materials.filter((m) => picked.includes(m));
  const pick = (m: string, more: boolean) =>
    setPicked(more ? (sel.includes(m) ? sel.filter((x) => x !== m) : [...sel, m]) : (sel.length === 1 && sel[0] === m ? [] : [m]));
  const takes = (p: ManualProcess) => !!drag?.some((m) => !p.materials.some((u) => u.material.toLowerCase() === m.toLowerCase()));
  const drop = (p: ManualProcess) => { if (drag) props.onChange(putMaterial(g, p.id, drag)); setDrag(null); setOver(null); };
  const isNew = editing && !g.processes.some((p) => p.id === editing.id);

  return (
    <div className="wk-man">
      <div className="wk-man-h">
        <span className="muted small">Entered here because BMES has no model group for this model. Save the model information to keep changes.</span>
        <span className="wk-man-hb">
          <button className="btn sm" disabled={pasting || !!editing} onClick={() => setPasting(true)}>Paste NG Rate header</button>
          {source && <button className="btn sm" onClick={copy}>Copy from Production Input {source.name} ({source.build})</button>}
        </span>
      </div>
      <AiRequest modelName={props.modelName} value={g} onApply={(next) => { props.onChange(next); setEditing(null); }} />
      {pasting && <SheetPaste value={g} onClose={() => setPasting(false)} onApply={(next) => { props.onChange(next); setPasting(false); setEditing(null); }} />}
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
                    <tr className={over === p.id ? "wk-man-over" : undefined}
                      onDragOver={(e) => { if (!takes(p)) return; e.preventDefault(); e.dataTransfer.dropEffect = "copy"; setOver(p.id); }}
                      onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setOver((o) => (o === p.id ? null : o)); }}
                      onDrop={(e) => { e.preventDefault(); drop(p); }}>
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
                <tr key={m} className={"wk-man-drag" + (sel.includes(m) ? " wk-man-picked" : "")} draggable
                  title="Drag onto a process to add it; Ctrl+click to pick several" onClick={(e) => pick(m, e.ctrlKey || e.metaKey)}
                  onDragStart={(e) => {
                    const list = sel.includes(m) ? sel : [m];
                    e.dataTransfer.setData("text/plain", list.join("\n")); e.dataTransfer.effectAllowed = "copy"; setDrag(list);
                  }}
                  onDragEnd={() => { setDrag(null); setOver(null); }}>
                  <td className="l wk-man-mat"><b>{m}</b></td>
                  <td className="muted">{usage(g, m)} proc</td>
                  <td className="act"><button className="wk-link bad" onClick={(e) => { e.stopPropagation(); removeMat(m); }} aria-label={`Remove ${m}`}>✕</button></td>
                </tr>
              ))}
              {!g.materials.length && <tr><td colSpan={3} className="wk-it-empty">No material yet.</td></tr>}
            </tbody>
          </table>
          {g.materials.length > 0 && (
            <div className="muted small wk-man-hint">
              {sel.length > 1 && <><b>{sel.length} picked</b> · <button className="wk-link" onClick={() => setPicked([])}>Clear</button><br /></>}
              Drag a material onto a process to add it; Ctrl+click to pick several and drag them together (1 per unit; change the qty with Edit).
            </div>
          )}
          <div className="wk-man-newmat">
            <input className="inp sm" value={newMat} placeholder="New material name" onChange={(e) => setNewMat(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") addMat(); }} />
            <button className="btn sm" disabled={!newMat.trim()} onClick={addMat}>+ Add</button>
          </div>
          <BomImport modelName={props.modelName} value={g} onApply={props.onChange} />
        </div>
      </div>
    </div>
  );
}
