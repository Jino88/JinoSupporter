import { useEffect, useState, type DragEvent } from "react";
import { workingApi } from "./api";
import { bmesApi } from "./bmes";
import {
  MANUAL_TYPES, addMaterial, addMaterials, aiAnswer, bomGlue, bomMaterials, blankProcess, deleteMaterial, deleteMaterials, deleteProcess, editProcess,
  glueOf, isBond, moveProcess, ngFromText, putMaterial, saveProcess, usage,
} from "./manual";
import type { ManualGroup as Group, ManualProcess, ManualType, ManualUse } from "./manual";
import { useTr } from "./i18n";
import { mergeSheet, parseNgHeader } from "./ngSheet";
import { loadInstructions } from "./workInstruction";

/** The header rows of a Daily NG Rate sheet pasted here: the processes and NG items read from them, put into the table. */
function SheetPaste(props: { value: Group; onApply: (g: Group) => void; onClose: () => void }) {
  const tr = useTr();
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
      <textarea className="inp sm wk-ta" rows={4} value={text} autoFocus placeholder={tr("Paste the header rows here (Ctrl+V)")}
        onChange={(e) => setText(e.target.value)} />
      {sheet?.problem && <div className="pi-dirty small">{sheet.problem}</div>}
      {merged && (
        <>
          <div className="small"><b>{sheet!.processes.length} {tr("processes")}</b>
            {merged.dropped.length > 0 && <span className="pi-dirty"> · {tr("removes")} {merged.dropped.join(", ")}</span>}
          </div>
          {sheet!.warnings.map((w) => <div key={w} className="pi-dirty small">{w}</div>)}
          <div className="pi-wrap wk-man-paste-list">
            <table className="wk-it">
              <thead><tr><th className="fit">#</th><th className="l">{tr("Process name")}</th><th className="l">{tr("NG items")}</th></tr></thead>
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
        <button className="btn sm" onClick={props.onClose}>{tr("Cancel")}</button>
        <button className="btn sm primary" disabled={!merged} onClick={apply}>{tr("Put into the table")}</button>
      </div>
    </div>
  );
}

/**
 * A request typed for AI (with cells copied from Excel if wanted), e.g. "fill the table from this header" or "put LOWER
 * into processes 3 to 6". The changed table it answers is shown first and goes into the table only when asked.
 */
function AiRequest(props: { modelName: string; value: Group; onApply: (g: Group) => void }) {
  const tr = useTr();
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
      <textarea className="inp sm wk-ta" rows={3} value={text} placeholder={tr("What should change in the table?")} disabled={!!busy}
        onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void send(); } }} />
      {error && <div className="pi-dirty small">{error}</div>}
      {answer && (
        <>
          <div className="small">
            {answer.note && <div>{answer.note}</div>}
            <b>{answer.group.processes.length} {tr("processes")}</b> · {answer.group.materials.length} {tr("materials")}
            {answer.added.length > 0 && <span> · {tr("adds")} {answer.added.join(", ")}</span>}
            {answer.removed.length > 0 && <span className="pi-dirty"> · {tr("removes")} {answer.removed.join(", ")}</span>}
            <span className="muted"> · {answer.model}</span>
          </div>
          <div className="pi-wrap wk-man-paste-list">
            <table className="wk-it">
              <thead><tr><th className="fit">#</th><th className="fit">{tr("Type")}</th><th className="l">{tr("Process name")}</th><th className="l">{tr("NG items")}</th><th className="l">{tr("Materials")}</th></tr></thead>
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
        {busy && <span className="muted small">{tr("Asking AI…")}</span>}
        {busy && <button className="btn sm" onClick={() => busy.abort()}>{tr("Stop")}</button>}
        {answer && <button className="btn sm" onClick={() => setAnswer(null)}>{tr("Discard")}</button>}
        {answer
          ? <button className="btn sm primary" onClick={apply}>{tr("Put into the table")}</button>
          : <button className="btn sm primary" disabled={!text.trim() || !!busy} onClick={() => void send()}>{tr("Ask AI")}</button>}
      </div>
    </div>
  );
}

/**
 * The product's direct materials (BOM level 1) put into the material list: read from the local BOM cache, or from BMES
 * when the cache has none or when asked. The new ones are listed with a check box first and added on Add.
 */
function BomImport(props: { modelName: string; value: Group; onApply: (g: Group) => void }) {
  const tr = useTr();
  const [product, setProduct] = useState(props.modelName);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [found, setFound] = useState<{ names: string[]; glue: string[]; level1: number; from: string } | null>(null);
  const [off, setOff] = useState<string[]>([]);
  const load = async (fromBmes: boolean) => {
    const name = product.trim();
    if (!name || busy) return;
    setBusy(true); setError(""); setFound(null); setOff([]);
    try {
      let bom = fromBmes ? null : await bmesApi.bom(name);
      if (!bom?.rows.length) bom = await bmesApi.fetchBom(name);
      if (!bom.rows.length) { setError(bom.error || `No BOM found for ${name}.`); return; }
      const names = bomMaterials(props.value, bom.rows);
      const glue = bomGlue(bom.rows);
      setFound({
        names, glue,
        level1: bom.rows.filter((r) => r.level === 1).length,
        from: bom.source === "server" ? "BMES" : "the local BOM cache",
      });
      // a bond (counted in G / ML) is not a material of the process table: left out unless ticked
      setOff(names.filter((n) => glue.includes(n)));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(false); }
  };
  const take = found ? found.names.filter((n) => !off.includes(n)) : [];
  const apply = () => { props.onApply(addMaterials(props.value, take)); setFound(null); };
  return (
    <div className="wk-man-bom">
      <div className="small"><b>{tr("Load from BOM")}</b> <span className="muted">({tr("direct materials, depth 1")})</span></div>
      <div className="wk-man-newmat">
        <input className="inp sm" value={product} placeholder={tr("Product name in BMES")} disabled={busy}
          onChange={(e) => setProduct(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") void load(false); }} />
        <button className="btn sm" disabled={!product.trim() || busy} onClick={() => void load(false)}>{busy ? tr("Loading…") : tr("Load")}</button>
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
              {found.glue.includes(n) && <span className="wk-man-bondtag">{tr("Bond")}</span>}
            </label>
          ))}
          <div className="wk-man-paste-b">
            <button className="btn sm" onClick={() => setFound(null)}>{tr("Cancel")}</button>
            {found.names.length > 0 && <button className="btn sm primary" disabled={!take.length} onClick={apply}>{tr("Add")} {take.length}</button>}
          </div>
        </>
      )}
    </div>
  );
}

/** A qty per unit typed in place: kept as typed while it is not a number above 0, put into the process when it is. */
function QtyCell(props: { value: number; onChange: (n: number) => void }) {
  const [text, setText] = useState(String(props.value));
  useEffect(() => { setText((cur) => (Number(cur) === props.value ? cur : String(props.value))); }, [props.value]);
  const ok = Number(text) > 0;
  return (
    <input className={"wk-man-in wk-man-qty" + (ok ? "" : " bad")} type="number" min={0} step="any" value={text} aria-label="Qty per unit"
      onChange={(e) => { setText(e.target.value); if (Number(e.target.value) > 0) props.onChange(Number(e.target.value)); }}
      onBlur={() => { if (!ok) setText(String(props.value)); }} />
  );
}

/**
 * One process changed in place, each change put into the table as it is made: type, names, NG items (one per line)
 * and the materials it uses per unit. Materials can also be dragged onto the row from the material list.
 */
function ProcessRow(props: {
  no: number; count: number; value: ManualProcess; materials: string[]; focus: boolean; over: boolean;
  /** Whether a material is a bond (marked so it is told from the parts). */
  bond: (material: string) => boolean;
  onChange: (p: ManualProcess) => void; onMove: (by: -1 | 1) => void; onDelete: () => void;
  onDragOver: (e: DragEvent) => void; onDragLeave: (e: DragEvent) => void; onDrop: () => void;
}) {
  const tr = useTr();
  const p = props.value;
  const set = (f: Partial<ManualProcess>) => props.onChange({ ...p, ...f });
  // the names and NG items stay as written in their boxes (that is what is saved); their translation is shown under them
  const shown = (text: string) => { const t = tr(text); return t !== text ? t : ""; };
  const nameTr = shown(p.name);
  const ngTr = p.ngItems.map(shown);
  // the NG items as typed (blank lines while typing); the process gets them cleaned
  const [ng, setNg] = useState(p.ngItems.join("\n"));
  const ngNow = p.ngItems.join("\n");
  useEffect(() => { setNg((cur) => (ngFromText(cur).join("\n") === ngNow ? cur : ngNow)); }, [ngNow]);
  const unused = props.materials.filter((m) => !p.materials.some((u) => u.material.toLowerCase() === m.toLowerCase()));
  const setUse = (i: number, f: Partial<ManualUse>) => set({ materials: p.materials.map((u, j) => (j === i ? { ...u, ...f } : u)) });
  const trim = (f: "name" | "localName") => { if (p[f] !== p[f].trim()) set({ [f]: p[f].trim() }); };
  return (
    <tr className={props.over ? "wk-man-over" : undefined} onDragOver={props.onDragOver} onDragLeave={props.onDragLeave}
      onDrop={(e) => { e.preventDefault(); props.onDrop(); }}>
      <td className="n">{props.no}</td>
      <td>
        <select className={"wk-man-in wk-man-type t-" + p.type.toLowerCase()} value={p.type} aria-label={tr("Type")}
          onChange={(e) => set({ type: e.target.value as ManualType })}>
          {MANUAL_TYPES.map((t) => <option key={t}>{t}</option>)}
        </select>
      </td>
      <td className="l">
        <input className={"wk-man-in wk-man-name" + (p.name.trim() ? "" : " bad")} value={p.name} placeholder={tr("Process name")} autoFocus={props.focus}
          onChange={(e) => set({ name: e.target.value })} onBlur={() => trim("name")} />
        {nameTr && <div className="wk-man-tr">{nameTr}</div>}
        <input className="wk-man-in wk-man-local" value={p.localName} placeholder={tr("Local name")}
          onChange={(e) => set({ localName: e.target.value })} onBlur={() => trim("localName")} />
      </td>
      <td className="l">
        <textarea className="wk-man-in wk-man-ng" rows={Math.max(1, ng.split("\n").length)} value={ng} placeholder={tr("NG items, one per line")}
          onChange={(e) => { setNg(e.target.value); set({ ngItems: ngFromText(e.target.value) }); }}
          onBlur={() => setNg(ngNow)} />
        {ngTr.some(Boolean) && <div className="wk-man-tr">{p.ngItems.map((n, i) => <div key={i}>{ngTr[i] || n}</div>)}</div>}
      </td>
      <td className="l">
        {p.materials.map((u, i) => (
          <div className={"wk-man-use" + (props.bond(u.material) ? " wk-man-bond" : "")} key={u.material + i}>
            <select className="wk-man-in wk-man-umat" value={u.material} aria-label="Material" onChange={(e) => setUse(i, { material: e.target.value })}>
              {[u.material, ...unused].filter(Boolean).map((m) => <option key={m}>{m}</option>)}
            </select>
            <span className="muted">×</span>
            <QtyCell value={u.qtyPerUnit} onChange={(n) => setUse(i, { qtyPerUnit: n })} />
            <button className="wk-link bad wk-man-x" onClick={() => set({ materials: p.materials.filter((_, j) => j !== i) })}
              aria-label={`Remove ${u.material}`}>✕</button>
          </div>
        ))}
        {unused.length > 0
          ? <button className="wk-link small wk-man-addmat" onClick={() => set({ materials: [...p.materials, { material: unused[0], qtyPerUnit: 1 }] })}>{tr("+ Material")}</button>
          : !props.materials.length && <span className="muted small">{tr("Add materials to the list first.")}</span>}
      </td>
      <td className="act">
        <button className="wk-link" disabled={props.no === 1} onClick={() => props.onMove(-1)} aria-label="Move up">↑</button>
        <button className="wk-link" disabled={props.no === props.count} onClick={() => props.onMove(1)} aria-label="Move down">↓</button>
        <button className="wk-link bad" onClick={props.onDelete}>{tr("Delete")}</button>
      </td>
    </tr>
  );
}

/**
 * Model group typed in Model Info for a model BMES > Make Model Group does not have: a table of the processes with
 * their NG items and the materials each uses per unit, and the material list. Kept with the model information (Save).
 */
export default function ManualGroup(props: {
  modelName: string;
  /** Product the BOM is read from first (the sub-group's name for a sub-group's table); the model name when unset. */
  product?: string;
  value: Group; onChange: (g: Group) => void; onError: (e: unknown) => void;
  /** Material keys of the glue the model's Work Instructions name; read from them when unset. */
  glue?: Set<string>;
}) {
  const tr = useTr();
  const g = props.value;
  // the process just added, whose name box gets the cursor
  const [added, setAdded] = useState<string | null>(null);
  const [newMat, setNewMat] = useState("");
  const [pasting, setPasting] = useState(false);
  // materials picked in the list (Ctrl+click for several), the ones dragged onto a process row, and the row they are over
  const [picked, setPicked] = useState<string[]>([]);
  const [drag, setDrag] = useState<string[] | null>(null);
  const [over, setOver] = useState<string | null>(null);
  const [readGlue, setReadGlue] = useState<Set<string>>(() => new Set());
  const glue = props.glue ?? readGlue;

  // the bonds the Work Instructions name (Keo · SL233LF), told from the parts in the lists
  useEffect(() => {
    if (props.glue) return;
    let live = true;
    loadInstructions(props.modelName).then((doc) => { if (live) setReadGlue(glueOf(doc)); }).catch(() => {});
    return () => { live = false; };
  }, [props.modelName, props.glue]);

  const addMat = () => {
    try { props.onChange(addMaterial(g, newMat)); setNewMat(""); } catch (e) { props.onError(e); }
  };
  const removeMat = (m: string) => {
    const n = usage(g, m);
    if (n && !window.confirm(`${m} is used in ${n} process${n > 1 ? "es" : ""}. Remove it from them too?`)) return;
    props.onChange(deleteMaterial(g, m));
  };
  const bond = (m: string) => isBond(m, glue);
  const bonds = g.materials.filter(bond);
  const removeBonds = () => {
    if (window.confirm(`Remove the ${bonds.length} bond${bonds.length > 1 ? "s" : ""} (${bonds.join(", ")}) from the material list and from every process?`))
      props.onChange(deleteMaterials(g, bonds));
  };
  const removeProc = (p: ManualProcess) => {
    if (window.confirm(`Delete process ${p.name}?`)) props.onChange(deleteProcess(g, p.id));
  };
  // a change is put into the table as typed; names are trimmed when the box is left
  const change = (p: ManualProcess) => props.onChange(editProcess(g, p));
  const add = () => {
    const p = blankProcess(g.processes.length ? g.processes[g.processes.length - 1].type : "SUB");
    props.onChange(saveProcess(g, p)); setAdded(p.id);
  };
  const sel = g.materials.filter((m) => picked.includes(m));
  const pick = (m: string, more: boolean) =>
    setPicked(more ? (sel.includes(m) ? sel.filter((x) => x !== m) : [...sel, m]) : (sel.length === 1 && sel[0] === m ? [] : [m]));
  const takes = (p: ManualProcess) => !!drag?.some((m) => !p.materials.some((u) => u.material.toLowerCase() === m.toLowerCase()));
  const drop = (p: ManualProcess) => { if (drag) props.onChange(putMaterial(g, p.id, drag)); setDrag(null); setOver(null); };

  const matRow = (m: string) => (
    <tr key={m} className={"wk-man-drag" + (sel.includes(m) ? " wk-man-picked" : "") + (bond(m) ? " wk-man-bond" : "")} draggable
      title={tr("Drag onto a process to add it; Ctrl+click to pick several")} onClick={(e) => pick(m, e.ctrlKey || e.metaKey)}
      onDragStart={(e) => {
        const list = sel.includes(m) ? sel : [m];
        e.dataTransfer.setData("text/plain", list.join("\n")); e.dataTransfer.effectAllowed = "copy"; setDrag(list);
      }}
      onDragEnd={() => { setDrag(null); setOver(null); }}>
      <td className="l wk-man-mat"><b>{m}</b>{bond(m) && <span className="wk-man-bondtag">{tr("Bond")}</span>}</td>
      <td className="muted">{usage(g, m)} {tr("proc")}</td>
      <td className="act"><button className="wk-link bad" onClick={(e) => { e.stopPropagation(); removeMat(m); }} aria-label={`Remove ${m}`}>✕</button></td>
    </tr>
  );

  return (
    <div className="wk-man">
      <div className="wk-man-h">
        <span className="wk-man-hb">
          <button className="btn sm" disabled={pasting} onClick={() => setPasting(true)}>{tr("Paste NG Rate header")}</button>        </span>
      </div>
      <AiRequest modelName={props.modelName} value={g} onApply={props.onChange} />
      {pasting && <SheetPaste value={g} onClose={() => setPasting(false)} onApply={(next) => { props.onChange(next); setPasting(false); }} />}
      <div className="wk-man-body">
        <div className="pi-wrap wk-man-procs">
          <table className="wk-it wk-man-t">
            <thead>
              <tr>
                <th className="fit">#</th><th className="fit">{tr("Type")}</th><th className="l">{tr("Process name")}</th>
                <th className="l">{tr("NG items")}</th><th className="l">{tr("Materials · qty per unit")}</th><th className="fit"></th>
              </tr>
            </thead>
            <tbody>
              {g.processes.map((p, i) => (
                <ProcessRow key={p.id} no={i + 1} count={g.processes.length} value={p} materials={g.materials} bond={bond}
                  focus={added === p.id} over={over === p.id}
                  onChange={change} onMove={(by) => props.onChange(moveProcess(g, p.id, by))} onDelete={() => removeProc(p)}
                  onDragOver={(e) => { if (!takes(p)) return; e.preventDefault(); e.dataTransfer.dropEffect = "copy"; setOver(p.id); }}
                  onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node | null)) setOver((o) => (o === p.id ? null : o)); }}
                  onDrop={() => drop(p)} />
              ))}
              {!g.processes.length && <tr><td colSpan={6} className="wk-it-empty">{tr("No process yet.")}</td></tr>}
            </tbody>
          </table>
          <button className="btn sm wk-man-add" onClick={add}>{tr("+ Add process")}</button>
        </div>
        <div className="wk-man-mats">
          <table className="wk-it">
            <thead><tr><th className="l">{tr("Material")}</th><th className="fit">{tr("Used in")}</th><th className="fit"></th></tr></thead>
            <tbody>
              {g.materials.filter((m) => !bond(m)).map(matRow)}
              {bonds.length > 0 && (
                <tr className="wk-man-bondh">
                  <td colSpan={3} className="l">
                    <b>{tr("Bond")}</b> <span className="muted">· {bonds.length} · {tr("glue, not a part")}</span>
                    <button className="wk-link bad wk-man-bondx" onClick={removeBonds}>{tr("Remove bonds")}</button>
                  </td>
                </tr>
              )}
              {bonds.map(matRow)}
              {!g.materials.length && <tr><td colSpan={3} className="wk-it-empty">{tr("No material yet.")}</td></tr>}
            </tbody>
          </table>
          {sel.length > 1 && (
            <div className="muted small wk-man-hint">
              <b>{sel.length} {tr("picked")}</b> · <button className="wk-link" onClick={() => setPicked([])}>{tr("Clear")}</button>
            </div>
          )}
          <div className="wk-man-newmat">
            <input className="inp sm" value={newMat} placeholder={tr("New material name")} onChange={(e) => setNewMat(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter") addMat(); }} />
            <button className="btn sm" disabled={!newMat.trim()} onClick={addMat}>{tr("+ Add")}</button>
          </div>
          <BomImport modelName={props.product ?? props.modelName} value={g} onApply={props.onChange} />
        </div>
      </div>
    </div>
  );
}
