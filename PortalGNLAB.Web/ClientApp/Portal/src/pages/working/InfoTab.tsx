import { useEffect, useState } from "react";
import { workingApi } from "./api";
import BmesPanel from "./BmesPanel";
import { PHASES, type ModelInfo, type WorkingModel } from "./types";

const FIELDS: { key: keyof ModelInfo; label: string; placeholder?: string }[] = [
  { key: "customer", label: "Customer" },
  { key: "build", label: "Build", placeholder: "PV2" },
  { key: "owner", label: "PM owner" },
  { key: "line", label: "Line" },
  { key: "sopTarget", label: "SOP target", placeholder: "yyyy-MM-dd" },
];

/** Model information form and the BMES model group of the model. */
export default function InfoTab(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
}) {
  const { model } = props;
  const [info, setInfo] = useState<ModelInfo>(model.info ?? {});
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState("");
  const dirty = JSON.stringify(info) !== JSON.stringify(model.info ?? {});
  useEffect(() => props.onDirty(dirty), [dirty]);

  const set = (key: keyof ModelInfo, value: string) => setInfo((cur) => ({ ...cur, [key]: value }));
  const save = async () => {
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, info, model.revision);
      setInfo(m.info ?? {});
      props.onSaved(m);
      setSaved("Saved.");
      window.setTimeout(() => setSaved(""), 2500);
    } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  const phaseIdx = PHASES.indexOf((info.phase ?? "") as (typeof PHASES)[number]);

  return (
    <>
      <section className="card pi-sec">
        <div className="pi-sec-h">
          <h2>Model information</h2>
          <div className="pi-legend">
            {dirty && <span className="pi-dirty">Unsaved</span>}
            {saved && <span className="pi-msg">{saved}</span>}
            {model.updatedAt && <span>Updated {model.updatedAt.slice(0, 16).replace("T", " ")} by {model.updatedBy || "—"}</span>}
            <button className="btn sm primary" disabled={!dirty || busy} onClick={save}>{busy ? "Saving…" : "Save"}</button>
          </div>
        </div>
        <div className="pi-sec-b">
          <div className="wk-phase-steps" role="radiogroup" aria-label="Phase">
            {PHASES.map((p, i) => (
              <button key={p} role="radio" aria-checked={info.phase === p}
                className={"wk-step" + (i < phaseIdx ? " done" : i === phaseIdx ? " now" : "")}
                onClick={() => set("phase", info.phase === p ? "" : p)}>
                <i>{i < phaseIdx ? "✓" : i + 1}</i>{p}
              </button>
            ))}
          </div>
          <div className="wk-form">
            <div className="wk-f"><label>Model</label><div className="wk-ro">{model.name}</div></div>
            {FIELDS.map((f) => (
              <div className="wk-f" key={f.key}>
                <label htmlFor={"wk-" + f.key}>{f.label}</label>
                <input id={"wk-" + f.key} className="inp" value={info[f.key] ?? ""} placeholder={f.placeholder}
                  onChange={(e) => set(f.key, e.target.value)} />
              </div>
            ))}
            <div className="wk-f wide">
              <label htmlFor="wk-note">Note</label>
              <textarea id="wk-note" className="inp wk-ta" rows={4} value={info.note ?? ""} onChange={(e) => set("note", e.target.value)} />
            </div>
          </div>
        </div>
      </section>

      <BmesPanel modelName={model.name} saved={info.bmesMaterial ?? ""} onError={props.onError}
        onPick={(m) => setInfo(({ bmesMaterial: _old, ...rest }) => (m ? { ...rest, bmesMaterial: m } : rest))} />
    </>
  );
}
