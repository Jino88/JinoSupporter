import { Fragment, useState } from "react";
import { NumInput, TextInput } from "./fields";
import {
  PROCESS_TYPES,
  SEGMENTS,
  SEGMENT_KEYS,
  addMaterial,
  addNgItem,
  addProcess,
  addStage,
  deleteMaterial,
  deleteNgItem,
  deleteProcess,
  deleteProcessMaterial,
  deleteStage,
  linkMaterial,
  moveProcess,
  renameNgItem,
  updateProcess,
  updateProcessMaterial,
  updateStage,
} from "./logic";
import type { ModelSetup, ProcessType, StageSource } from "./types";

const DRAG_TYPE = "text/x-pi-material";

export const BUILDS = ["PV2", "PV1", "DV", "LPR", "MP"];

const SOURCES: [StageSource | "", string][] = [
  ["", "Typed in the plan"],
  ["sub", "Sub Dry in (segment A output)"],
  ["main", "Main Dry in (segment C output)"],
  ["func", "Last function test output"],
  ["vis", "Packing input"],
  ["oqc", "Send to OQC"],
];

export default function SetupTab(props: {
  name: string;
  build: string;
  setup: ModelSetup;
  onSetup: (setup: ModelSetup) => void;
  onModel: (change: { name?: string; build?: string }) => void;
  onDelete: () => void;
  onError: (message: string) => void;
}) {
  const { setup, onSetup } = props;
  const [newMat, setNewMat] = useState("");
  const [dropOn, setDropOn] = useState<string | null>(null);
  const procs = setup.processes;

  function addMat() {
    try {
      onSetup(addMaterial(setup, newMat));
      setNewMat("");
    } catch (e) {
      props.onError((e as Error).message);
    }
  }

  return (
    <>
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Model</h2>
          <button className="btn sm danger" onClick={props.onDelete}>Delete model</button>
        </div>
        <div className="pi-sec-b form-row">
          <label>Model name<TextInput className="name" value={props.name} label="Model name"
            onCommit={(v) => v.trim() && props.onModel({ name: v.trim() })} /></label>
          <label>Build
            <select className="pi-sel" value={props.build} onChange={(e) => props.onModel({ build: e.target.value })}>
              {[...new Set([props.build, ...BUILDS])].map((b) => <option key={b}>{b}</option>)}
            </select>
          </label>
        </div>
      </section>

      <div className="pi-setup">
        <section className="card pi-sec">
          <div className="pi-sec-h">
            <h2>Processes, NG items and materials per unit <span className="cnt">Step 3</span></h2>
            <span className="muted small">Numbers renumber automatically. ↑ / ↓ at the edge of a block moves the process into the next block. Material qty = pieces used per 1 unit going into the process.</span>
          </div>
          <div className="pi-wrap">
            <table className="pi-pe">
              <thead>
                <tr><th>#</th><th>Process Type</th><th>Process name</th><th>NG items (listed down)</th><th>Materials used · qty per unit</th><th></th></tr>
              </thead>
              <tbody>
                {SEGMENT_KEYS.map((sk) => (
                  <Fragment key={sk}>
                    <tr className="seg-row">
                      <td colSpan={6}>
                        <span className="bar" style={{ background: SEGMENTS[sk].color }} />
                        {SEGMENTS[sk].name}<span className="why">{SEGMENTS[sk].why}</span>
                        <button className="pi-lk" onClick={() => onSetup(addProcess(setup, sk))}>+ Add process</button>
                      </td>
                    </tr>
                    {procs.map((p, idx) => p.segment !== sk ? null : (
                      <tr className={"plast" + (dropOn === p.id ? " pi-drop" : "")} key={p.id} data-pid={p.id}
                        onDragOver={(e) => { if (e.dataTransfer.types.includes(DRAG_TYPE)) { e.preventDefault(); setDropOn(p.id); } }}
                        onDragLeave={() => setDropOn((cur) => (cur === p.id ? null : cur))}
                        onDrop={(e) => {
                          const name = e.dataTransfer.getData(DRAG_TYPE);
                          setDropOn(null);
                          if (!name) return;
                          e.preventDefault();
                          if (p.materials.some((m) => m.material === name)) props.onError(`${name} is already used by ${p.name}.`);
                          else onSetup(linkMaterial(setup, p.id, name));
                        }}>
                        <td className="pi-pno rs">{idx + 1}</td>
                        <td className="rs">
                          <select className="pi-sel" value={p.type} aria-label="Process type"
                            title="FUNCTION counts in the function NG rate, the others in the process NG rate"
                            onChange={(e) => onSetup(updateProcess(setup, p.id, { type: e.target.value as ProcessType }))}>
                            {PROCESS_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                          </select>
                        </td>
                        <td className="rs">
                          <TextInput className="name" value={p.name} label="Process name"
                            onCommit={(v) => onSetup(updateProcess(setup, p.id, { name: v.trim() || "Unnamed" }))} />
                          <TextInput className="vi" value={p.localName} placeholder="Local name (optional)" label="Local name"
                            onCommit={(v) => onSetup(updateProcess(setup, p.id, { localName: v.trim() }))} />
                        </td>
                        <td className="rs">
                          <div className="pi-vl">
                            {p.ngItems.map((item) => (
                              <div className="r" key={item.id}>
                                <TextInput className="dname" value={item.name} label="NG item"
                                  onCommit={(v) => onSetup(renameNgItem(setup, p.id, item.id, v))} />
                                <button className="pi-x" title="Delete NG item" aria-label={`Delete NG ${item.name}`}
                                  onClick={() => onSetup(deleteNgItem(setup, p.id, item.id))}>×</button>
                              </div>
                            ))}
                            <button className="pi-lk" onClick={() => onSetup(addNgItem(setup, p.id))}>+ Add NG</button>
                          </div>
                        </td>
                        <td className="rs">
                          <div className="pi-vl">
                            {p.materials.length === 0 && <span className="muted small">Drop a material here</span>}
                            {p.materials.map((m, i) => (
                              <div className="r" key={i}>
                                <select className="pi-sel" value={m.material} aria-label="Material"
                                  onChange={(e) => onSetup(updateProcessMaterial(setup, p.id, i, { material: e.target.value }))}>
                                  {[...new Set([m.material, ...setup.materials])].map((n) => <option key={n}>{n}</option>)}
                                </select>
                                <NumInput decimal value={m.qtyPerUnit} label="Qty per unit"
                                  onCommit={(v) => onSetup(updateProcessMaterial(setup, p.id, i, { qtyPerUnit: v ?? 0 }))} />
                                <span className="muted small">/ unit</span>
                                <button className="pi-x" title="Remove material" aria-label="Remove material"
                                  onClick={() => onSetup(deleteProcessMaterial(setup, p.id, i))}>×</button>
                              </div>
                            ))}
                          </div>
                        </td>
                        <td className="rs" style={{ whiteSpace: "nowrap" }}>
                          <button className="pi-mv" title="Move up" onClick={() => onSetup(moveProcess(setup, p.id, -1))}>↑</button>{" "}
                          <button className="pi-mv" title="Move down" onClick={() => onSetup(moveProcess(setup, p.id, 1))}>↓</button>{" "}
                          <button className="pi-mv del" onClick={() => onSetup(deleteProcess(setup, p.id))}>Delete</button>
                        </td>
                      </tr>
                    ))}
                  </Fragment>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        <aside className="card pi-sec pi-matpanel">
          <div className="pi-sec-h">
            <h2>Materials <span className="cnt">Step 2</span></h2>
          </div>
          <div className="pi-sec-b">
            <div className="pi-hint">Drag a material onto a process row to use it there (1 per unit; change the qty in the row).</div>
            <div className="pi-matlist">
              {setup.materials.length === 0 && <div className="muted small">No material yet.</div>}
              {setup.materials.map((m) => (
                <div className="pi-matitem" key={m} draggable title="Drag onto a process"
                  onDragStart={(e) => { e.dataTransfer.setData(DRAG_TYPE, m); e.dataTransfer.effectAllowed = "copy"; }}
                  onDragEnd={() => setDropOn(null)}>
                  <span className="pi-grip" aria-hidden="true">⋮⋮</span>
                  <b>{m}</b>
                  <small>{procs.filter((p) => p.materials.some((x) => x.material === m)).length} proc</small>
                  <button className="pi-x" title="Remove material" aria-label={`Remove ${m}`}
                    onClick={() => onSetup(deleteMaterial(setup, m))}>×</button>
                </div>
              ))}
            </div>
            <div className="pi-matadd">
              <input className="pi-ti" placeholder="New material name" value={newMat} aria-label="New material name"
                onChange={(e) => setNewMat(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addMat()} />
              <button className="pi-lk" onClick={addMat}>+ Add material</button>
            </div>
          </div>
        </aside>
      </div>

      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Plan stages</h2>
          <span className="muted small">Rows of the production plan. A stage with a source takes its actual from the daily input.</span>
        </div>
        <div className="pi-sec-b">
          <div className="pi-vl">
            {setup.stages.map((s) => (
              <div className="r" key={s.id}>
                <TextInput className="name" value={s.name} label="Stage name"
                  onCommit={(v) => onSetup(updateStage(setup, s.id, { name: v.trim() || "Stage" }))} />
                <select className="pi-sel" value={s.source ?? ""} aria-label="Actual source"
                  onChange={(e) => onSetup(updateStage(setup, s.id, { source: (e.target.value || null) as StageSource | null }))}>
                  {SOURCES.map(([v, label]) => <option key={v} value={v}>{label}</option>)}
                </select>
                <button className="pi-x" title="Delete stage" aria-label={`Delete stage ${s.name}`}
                  onClick={() => onSetup(deleteStage(setup, s.id))}>×</button>
              </div>
            ))}
            <button className="pi-lk" onClick={() => onSetup(addStage(setup))}>+ Add stage</button>
          </div>
        </div>
      </section>
    </>
  );
}
