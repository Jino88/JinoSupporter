import { useEffect, useState } from "react";
import { workingApi } from "./api";
import { ProcessTables } from "./BmesPanel";
import { readManual, readSubManual, type ManualGroup } from "./manual";
import { readTempGroup } from "./tempGroup";
import type { WorkingModel } from "./types";

/**
 * Process tab: the processes, NG items and materials of the model (info.manualGroup, and info.subManual per sub-group of
 * the temporary model group, which is only read here). Saved with the model information.
 */
export default function ProcessTab(props: {
  model: WorkingModel;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
}) {
  const { model } = props;
  const read = (m: WorkingModel) => ({ manual: readManual(m.info.manualGroup), sub: readSubManual(m.info.subManual) });
  const [value, setValue] = useState(read(model));
  const [busy, setBusy] = useState(false);
  const [saved, setSaved] = useState("");
  const dirty = JSON.stringify(value) !== JSON.stringify(read(model));
  useEffect(() => props.onDirty(dirty), [dirty]);

  const save = async () => {
    setBusy(true);
    try {
      const m = await workingApi.saveModel(model.name, { ...model.info, manualGroup: value.manual, subManual: value.sub }, model.revision);
      setValue(read(m));
      props.onSaved(m);
      setSaved("Saved.");
      window.setTimeout(() => setSaved(""), 2500);
    } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  return (
    <section className="card pi-sec wk-info-sec">
      <div className="pi-sec-h">
        <h2>Process</h2>
        <div className="pi-legend">
          {dirty && <span className="pi-dirty">Unsaved</span>}
          {saved && <span className="pi-msg">{saved}</span>}
          {model.updatedAt && <span>Updated {model.updatedAt.slice(0, 16).replace("T", " ")} by {model.updatedBy || "—"}</span>}
          <button className="btn sm primary" disabled={!dirty || busy} onClick={save}>{busy ? "Saving…" : "Save"}</button>
        </div>
      </div>
      <div className="pi-sec-b">
        <ProcessTables modelName={model.name} manual={value.manual} onManual={(g: ManualGroup) => setValue((v) => ({ ...v, manual: g }))}
          temp={readTempGroup(model.info.tempGroup)} subManual={value.sub} onSubManual={(t) => setValue((v) => ({ ...v, sub: t }))}
          onError={props.onError} />
      </div>
    </section>
  );
}
