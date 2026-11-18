import { userName } from "../../userNames";
import { useEffect, useState } from "react";
import { workingApi } from "./api";
import { bmesApi, matchMaterial, type BmesMaterial } from "./bmes";
import { useTr } from "./i18n";
import { MANUAL } from "./manual";
import {
  MOLD_KINDS, blankMaterial, bomMaterialItems, cleanMaterials, materialGroups, materialProblem, moldCounts, readMaterials,
  type GroupDef, type MaterialItem, type MaterialList, type MoldKind,
} from "./materialList";
import { readTempGroup } from "./tempGroup";
import type { WorkingModel } from "./types";

/**
 * The materials of the group's BOM not in the list yet, read from the local BOM cache or from BMES, ticked to add
 * (direct materials, or every level when asked).
 */
function BomLoad(props: { group: GroupDef; rows: MaterialItem[]; onAdd: (rows: MaterialItem[]) => void }) {
  const tr = useTr();
  const [product, setProduct] = useState(props.group.product);
  const [all, setAll] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [found, setFound] = useState<{ items: MaterialItem[]; from: string } | null>(null);
  const [off, setOff] = useState<string[]>([]);
  const load = async (fromBmes: boolean) => {
    const name = product.trim();
    if (!name || busy) return;
    setBusy(true); setError(""); setFound(null); setOff([]);
    try {
      let bom = fromBmes ? null : await bmesApi.bom(name);
      if (!bom?.rows.length) bom = await bmesApi.fetchBom(name);
      if (!bom.rows.length) { setError(bom.error || `No BOM found for ${name}.`); return; }
      setFound({ items: bomMaterialItems(props.rows, bom.rows, all), from: bom.source === "server" ? "BMES" : "the local BOM cache" });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(false); }
  };
  const take = found ? found.items.filter((m) => !off.includes(m.id)) : [];
  return (
    <div className="wk-mat-bom">
      <div className="wk-man-newmat">
        <span className="small"><b>{tr("Load from BOM")}</b></span>
        <input className="inp sm" value={product} placeholder={tr("Product name in BMES")} disabled={busy} aria-label="BOM product"
          onChange={(e) => setProduct(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") void load(false); }} />
        <label className="small"><input type="checkbox" checked={all} onChange={(e) => setAll(e.target.checked)} /> {tr("Every level")}</label>
        <button className="btn sm" disabled={!product.trim() || busy} onClick={() => void load(false)}>{busy ? tr("Loading…") : tr("Load")}</button>
      </div>
      {error && <div className="pi-dirty small">{error}</div>}
      {found && (
        <>
          <div className="muted small">
            {found.items.length ? `${found.items.length} not in the list` : "All already in the list"} · from {found.from}
            {found.from !== "BMES" && <> · <button className="wk-link" onClick={() => void load(true)}>Read from BMES again</button></>}
          </div>
          {found.items.map((m) => (
            <label key={m.id} className="wk-man-bomi small">
              <input type="checkbox" checked={!off.includes(m.id)} onChange={(e) => setOff(e.target.checked ? off.filter((x) => x !== m.id) : [...off, m.id])} />
              <span className="mono">{m.code}</span> {m.name}{m.qty && <span className="muted"> · {m.qty} {m.unit}</span>}
            </label>
          ))}
          <div className="wk-man-paste-b">
            <button className="btn sm" onClick={() => setFound(null)}>{tr("Cancel")}</button>
            {found.items.length > 0 && (
              <button className="btn sm primary" disabled={!take.length} onClick={() => { props.onAdd(take); setFound(null); }}>{tr("Add")} {take.length}</button>
            )}
          </div>
        </>
      )}
    </div>
  );
}

/** The material table of one group, edited in place. */
export function MaterialTable(props: { rows: MaterialItem[]; onChange: (rows: MaterialItem[]) => void }) {
  const tr = useTr();
  const { rows } = props;
  const set = (id: string, patch: Partial<MaterialItem>) => props.onChange(rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  return (
    <table className="wk-mat-t">
      <thead>
        <tr>
          <th className="n">No</th><th>{tr("Code")}</th><th>{tr("Material name")}</th><th className="q">{tr("Qty")}</th><th>{tr("Unit")}</th>
          <th>{tr("Mold")}</th><th>{tr("Note")}</th><th></th>
        </tr>
      </thead>
      <tbody>
        {rows.length === 0 && <tr><td colSpan={8} className="muted small">{tr("No material yet. Add one or load them from the BOM.")}</td></tr>}
        {rows.map((r, i) => (
          <tr key={r.id} className={r.mold ? "mold-" + r.mold : undefined}>
            <td className="n">{i + 1}</td>
            <td><input className="inp sm mono" value={r.code} aria-label="Material code" onChange={(e) => set(r.id, { code: e.target.value })} /></td>
            <td><input className="inp sm" value={r.name} aria-label="Material name" onChange={(e) => set(r.id, { name: e.target.value })} /></td>
            <td className="q"><input className="inp sm" value={r.qty} aria-label="Qty" onChange={(e) => set(r.id, { qty: e.target.value })} /></td>
            <td><input className="inp sm wk-mat-unit" value={r.unit} aria-label="Unit" onChange={(e) => set(r.id, { unit: e.target.value })} /></td>
            <td>
              <select className="inp sm" value={r.mold} aria-label="Mold" onChange={(e) => set(r.id, { mold: e.target.value as MoldKind })}>
                {MOLD_KINDS.map((k) => <option key={k.key} value={k.key}>{tr(k.label)}</option>)}
              </select>
            </td>
            <td><input className="inp sm" value={r.note} aria-label="Note" onChange={(e) => set(r.id, { note: e.target.value })} /></td>
            <td><button className="wk-link danger" onClick={() => props.onChange(rows.filter((x) => x.id !== r.id))}>{tr("Remove")}</button></td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

/**
 * Material tab: the materials of each group of the model (product of its BMES model group, sub-group of its temporary
 * group, or the model), typed by hand or loaded from the BOM, each marked as made in a mass production or a sample mold.
 * Saved with the model information (info.materials).
 */
export default function MaterialTab(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
  /** BMES model groups already read (tests); else read here. */
  initial?: BmesMaterial[];
}) {
  const { model } = props;
  const tr = useTr();
  const [materials, setMaterials] = useState<BmesMaterial[] | null>(props.initial ?? null);
  const [value, setValue] = useState<MaterialList>(readMaterials(model.info.materials));
  const [picked, setPicked] = useState("");
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState("");
  const dirty = JSON.stringify(cleanMaterials(value)) !== JSON.stringify(cleanMaterials(readMaterials(model.info.materials)));
  useEffect(() => props.onDirty(dirty), [dirty]);

  useEffect(() => {
    if (props.initial) return;
    let live = true;
    // without BMES the groups are the temporary group's or the model itself
    bmesApi.materials().then((m) => { if (live) setMaterials(m); }).catch(() => { if (live) setMaterials([]); });
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const saved0 = model.info.bmesMaterial ?? "";
  const bmes = materials && saved0 !== MANUAL ? matchMaterial(model.name, materials, saved0).material : null;
  const groups = materialGroups(model.name, bmes, readTempGroup(model.info.tempGroup), value);
  const group = groups.find((g) => g.key === picked) ?? groups[0];
  const rows = value[group.key]?.rows ?? [];
  const setRows = (r: MaterialItem[]) => setValue((v) => ({ ...v, [group.key]: { label: group.label, rows: r } }));
  const problems = groups.map((g) => ({ g, p: materialProblem(value[g.key]?.rows ?? []) })).filter((x) => x.p);
  const counts = moldCounts(rows);

  const save = async () => {
    if (problems.length) return;
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, { ...model.info, materials: cleanMaterials(value) }, model.revision);
      setValue(readMaterials(m.info.materials));
      props.onSaved(m);
      setSaved(tr("Saved."));
      window.setTimeout(() => setSaved(""), 2500);
    } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr("Material")}</h2>
        <div className="pi-legend">
          {problems.length > 0 && <span className="pi-dirty">{problems[0].g.label}: {tr(problems[0].p)}</span>}
          {dirty && <span className="pi-dirty">{tr("Unsaved")}</span>}
          {saved && <span className="pi-msg">{saved}</span>}
          {model.updatedAt && <span>{tr("Updated")} {model.updatedAt.slice(0, 16).replace("T", " ")} · {userName(model.updatedBy) || "—"}</span>}
          <button className="btn sm primary" disabled={!dirty || busy || problems.length > 0} onClick={save}>{busy ? tr("Saving…") : tr("Save")}</button>
        </div>
      </div>
      <div className="pi-sec-b">
        {materials === null ? <div className="muted">Loading the BMES model groups…</div> : (
          <>
            <div className="wk-tg-tables" role="tablist" aria-label="Model groups">
              <span className="muted small">{tr("Materials of")}</span>
              {groups.map((g) => {
                const n = value[g.key]?.rows.length ?? 0;
                return (
                  <button key={g.key} role="tab" aria-selected={g === group} className={"wk-chip" + (g === group ? " on" : "")} onClick={() => setPicked(g.key)}>
                    {g.label}{n > 0 && <span className="cnt">{n}</span>}
                  </button>
                );
              })}
            </div>
            <div className="wk-mat-sum small">
              <span className="wk-mat-kind mass">{tr("Mass production mold")} {counts.mass}</span>
              <span className="wk-mat-kind sample">{tr("Sample mold")} {counts.sample}</span>
              {rows.length - counts.mass - counts.sample > 0 && <span className="muted">{tr("Not set")} {rows.length - counts.mass - counts.sample}</span>}
            </div>
            <MaterialTable rows={rows} onChange={setRows} />
            <div className="wk-mat-adds">
              <button className="btn sm" onClick={() => setRows([...rows, blankMaterial()])}>+ {tr("Material")}</button>
            </div>
            <BomLoad key={group.key} group={group} rows={rows} onAdd={(add) => setRows([...rows, ...add])} />
          </>
        )}
      </div>
    </section>
  );
}
