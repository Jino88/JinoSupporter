import { useEffect, useState } from "react";
import { workingApi } from "./api";
import BmesPanel from "./BmesPanel";
import { readManual, readSubManual } from "./manual";
import { readTempGroup } from "./tempGroup";
import { PHASES, phaseName, type ModelInfo, type WorkingModel } from "./types";

/** Model info with an old phase name (Proto, EVT …) shown as the current one; saving writes the new name. */
const current = (info: ModelInfo | undefined): ModelInfo => (info?.phase ? { ...info, phase: phaseName(info.phase) } : info ?? {});

/**
 * Model information: the phase and the BMES model group of the model. The stage (build), customer, PM owner, line,
 * SOP target and note are no longer edited here but are kept as they are saved.
 */
export default function InfoTab(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
}) {
  const { model } = props;
  const [info, setInfo] = useState<ModelInfo>(current(model.info));
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState("");
  const dirty = JSON.stringify(info) !== JSON.stringify(current(model.info));
  useEffect(() => props.onDirty(dirty), [dirty]);

  const set = (key: keyof ModelInfo, value: string) => setInfo((cur) => ({ ...cur, [key]: value }));
  const save = async () => {
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, info, model.revision);
      setInfo(current(m.info));
      props.onSaved(m);
      setSaved("Saved.");
      window.setTimeout(() => setSaved(""), 2500);
    } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  const phaseIdx = PHASES.indexOf((info.phase ?? "") as (typeof PHASES)[number]);

  return (
    <>
      <section className="card pi-sec wk-info-sec">
        <div className="pi-sec-h">
          <div className="wk-info-h">
            <h2>Model information</h2>
            <div className="wk-phase-steps" role="radiogroup" aria-label="Phase">
              {PHASES.map((p, i) => (
                <button key={p} role="radio" aria-checked={info.phase === p}
                  className={"wk-step" + (i < phaseIdx ? " done" : i === phaseIdx ? " now" : "")}
                  onClick={() => set("phase", info.phase === p ? "" : p)}>
                  <i>{i < phaseIdx ? "✓" : i + 1}</i>{p}
                </button>
              ))}
            </div>
          </div>
          <div className="pi-legend">
            {dirty && <span className="pi-dirty">Unsaved</span>}
            {saved && <span className="pi-msg">{saved}</span>}
            {model.updatedAt && <span>Updated {model.updatedAt.slice(0, 16).replace("T", " ")} by {model.updatedBy || "—"}</span>}
            <button className="btn sm primary" disabled={!dirty || busy} onClick={save}>{busy ? "Saving…" : "Save"}</button>
          </div>
        </div>
        <div className="pi-sec-b">
          <BmesPanel modelName={model.name} saved={info.bmesMaterial ?? ""} onError={props.onError}
            manual={readManual(info.manualGroup)} onManual={(g) => setInfo((cur) => ({ ...cur, manualGroup: g }))}
            temp={readTempGroup(info.tempGroup)} onTemp={(g) => setInfo((cur) => ({ ...cur, tempGroup: g }))}
            subManual={readSubManual(info.subManual)} onSubManual={(t) => setInfo((cur) => ({ ...cur, subManual: t }))}
            onPick={(m) => setInfo(({ bmesMaterial: _old, ...rest }) => (m ? { ...rest, bmesMaterial: m } : rest))} />
        </div>
      </section>
    </>
  );
}
