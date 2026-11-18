import { useEffect, useState } from "react";
import { useTr } from "./i18n";
import { MOLD_KINDS, materialKey, materialOptions, type MaterialList } from "./materialList";
import type { TaskMaterial } from "./plan";

/** The materials picked: in the order of the list, then the ones picked before that are no longer in it. */
export function pickedMaterials(list: MaterialList, keys: string[], before: TaskMaterial[]): TaskMaterial[] {
  const options = materialOptions(list);
  const listed = options.filter((o) => keys.includes(o.key)).map((o) => ({ code: o.item.code.trim(), name: o.item.name.trim() }));
  const gone = before.filter((m) => keys.includes(materialKey(m)) && !options.some((o) => o.key === materialKey(m)));
  return [...listed, ...gone];
}

/**
 * Popup to pick the materials of a material task from the model's material list (Model Info › Material), with their
 * mold; a material picked before that is no longer in the list stays until it is unticked.
 */
export default function MaterialPick(props: {
  title: string;
  materials: MaterialList;
  value: TaskMaterial[];
  onPick: (list: TaskMaterial[]) => void;
  onClose: () => void;
}) {
  const tr = useTr();
  const options = materialOptions(props.materials);
  const gone = props.value.filter((m) => !options.some((o) => o.key === materialKey(m)));
  const [keys, setKeys] = useState(props.value.map(materialKey));
  const [q, setQ] = useState("");
  const find = q.trim().toLowerCase();
  const shown = find ? options.filter((o) => `${o.item.code} ${o.item.name} ${o.groups.join(" ")}`.toLowerCase().includes(find)) : options;
  const tick = (key: string, on: boolean) => setKeys((k) => (on ? [...k, key] : k.filter((x) => x !== key)));
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") props.onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  return (
    <div className="modal-back" onClick={(e) => { if (e.target === e.currentTarget) props.onClose(); }}>
      <div className="modal wk-mpick" role="dialog" aria-modal="true" aria-label="Pick materials">
        <div className="wk-ai-pop-h">
          <h2>{tr("Pick materials")} · {props.title}</h2>
          <button className="btn sm" aria-label="Close" onClick={props.onClose}>×</button>
        </div>
        {options.length === 0 && gone.length === 0
          ? <div className="muted small">{tr("No material in Model Info › Material yet. Add them there (by hand or from the BOM) first.")}</div>
          : (
            <>
              <input className="inp sm wk-mpick-q" value={q} placeholder={tr("Search code, name or group")} aria-label="Search materials" onChange={(e) => setQ(e.target.value)} />
              <table className="wk-mat-t wk-mpick-t">
                <thead>
                  <tr><th></th><th>{tr("Code")}</th><th>{tr("Material name")}</th><th className="q">{tr("Qty")}</th><th>{tr("Mold")}</th><th>{tr("Group")}</th></tr>
                </thead>
                <tbody>
                  {shown.map((o) => {
                    const mold = MOLD_KINDS.find((k) => k.key === o.item.mold && k.key);
                    return (
                      <tr key={o.key} className={o.item.mold ? "mold-" + o.item.mold : undefined}>
                        <td><input type="checkbox" checked={keys.includes(o.key)} aria-label={`Pick ${o.item.code || o.item.name}`} onChange={(e) => tick(o.key, e.target.checked)} /></td>
                        <td className="mono">{o.item.code}</td>
                        <td>{o.item.name}</td>
                        <td className="q">{o.item.qty} {o.item.unit}</td>
                        <td>{mold ? <span className={"wk-mat-kind " + mold.key}>{tr(mold.label)}</span> : "—"}</td>
                        <td className="muted small">{o.groups.join(", ")}</td>
                      </tr>
                    );
                  })}
                  {gone.map((m) => (
                    <tr key={materialKey(m)} className="gone">
                      <td><input type="checkbox" checked={keys.includes(materialKey(m))} aria-label={`Pick ${m.code || m.name}`} onChange={(e) => tick(materialKey(m), e.target.checked)} /></td>
                      <td className="mono">{m.code}</td>
                      <td>{m.name}</td>
                      <td colSpan={3} className="muted small">{tr("No longer in the material list")}</td>
                    </tr>
                  ))}
                  {shown.length === 0 && gone.length === 0 && <tr><td colSpan={6} className="muted small">{tr("No material matches.")}</td></tr>}
                </tbody>
              </table>
            </>
          )}
        <div className="modal-actions">
          <button className="btn sm" onClick={props.onClose}>{tr("Cancel")}</button>
          <button className="btn sm primary" onClick={() => props.onPick(pickedMaterials(props.materials, keys, props.value))}>{tr("Select")} {keys.length}</button>
        </div>
      </div>
    </div>
  );
}
