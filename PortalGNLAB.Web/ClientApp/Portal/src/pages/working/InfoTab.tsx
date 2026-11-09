import { useEffect, useState } from "react";
import { workingApi } from "./api";
import { countByKind, isLate, isOpenIssue } from "./logic";
import { PHASES, type Entry, type ModelInfo, type WorkingModel } from "./types";

const FIELDS: { key: keyof ModelInfo; label: string; placeholder?: string }[] = [
  { key: "customer", label: "Customer" },
  { key: "build", label: "Build", placeholder: "PV2" },
  { key: "owner", label: "PM owner" },
  { key: "line", label: "Line" },
  { key: "sopTarget", label: "SOP target", placeholder: "yyyy-MM-dd" },
];

/** Model information form plus a short summary of the history counts. */
export default function InfoTab(props: {
  model: WorkingModel;
  entries: Entry[];
  today: string;
  onDirty: (dirty: boolean) => void;
  onSaved: (model: WorkingModel) => void;
  onError: (e: unknown) => void;
}) {
  const { model, entries, today } = props;
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

  const c = countByKind(entries);
  const open = entries.filter(isOpenIssue);
  const late = open.filter((e) => isLate(e, today)).length;
  const lastReview = entries.find((e) => e.kind === "review");
  const phaseIdx = PHASES.indexOf((info.phase ?? "") as (typeof PHASES)[number]);

  return (
    <>
      <section className="card pi-kpis wk-kpis">
        <div className="pi-kp"><div className="pi-kp-k">Phase</div><div className="pi-kp-v">{info.phase || "—"}</div><div className="pi-kp-n">{info.build || "No build"}</div></div>
        <div className="pi-kp"><div className="pi-kp-k">Open issues</div><div className={"pi-kp-v" + (open.length ? " bad" : "")}>{open.length}</div><div className="pi-kp-n">{late ? `${late} past due` : `${c.issue} in total`}</div></div>
        <div className="pi-kp"><div className="pi-kp-k">Dev history</div><div className="pi-kp-v">{c.dev}</div><div className="pi-kp-n">entries</div></div>
        <div className="pi-kp"><div className="pi-kp-k">Last review</div><div className="pi-kp-v">{lastReview?.result || "—"}</div><div className="pi-kp-n">{lastReview ? lastReview.date : "No review yet"}</div></div>
      </section>

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
    </>
  );
}
