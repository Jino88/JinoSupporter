import { useRef, useState } from "react";
import WorkInstructionTab from "./WorkInstructionTab";
import { workingApi } from "./api";
import { loadInstructions } from "./workInstruction";
import { readManual } from "./manual";
import { processListApi } from "./processList";
import {
  NOTHING_TO_DO, errorText, isWorkbook, jigSummary, manualAiFromInstructions, processSummary, runJig, runProcess, type DocsResult,
} from "./docsAi";
import type { WorkingModel } from "./types";
import "./modelDocs.css";

/** Model Info › Documents: Work Instructions, a JIG / Machine / Tester workbook, and the AI that makes Process and JIG/MC List from them. */
export default function ModelDocs(props: {
  model: WorkingModel; onSaved: (m: WorkingModel) => void; onOpen: (k: "process" | "jigmc") => void; onError: (e: unknown) => void;
}) {
  const { model } = props;
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);
  const [result, setResult] = useState<DocsResult | null>(null);
  const pick = useRef<HTMLInputElement>(null);

  const run = async () => {
    if (busy) return;
    setResult(null);
    let hasInstructions = false;
    try { hasInstructions = (await loadInstructions(model.name)) !== null; } catch (e) { props.onError(e); return; }
    if (!hasInstructions && !file) { setStatus(NOTHING_TO_DO); return; }
    setBusy(true);
    const out: DocsResult = {};
    if (hasInstructions) {
      setStatus("Process: the AI is reading the Work Instructions…");
      try {
        out.process = await runProcess({
          hasInstructions,
          currentGroup: readManual(model.info.manualGroup),
          ask: (g) => manualAiFromInstructions(model.name, g),
          confirm: (t) => window.confirm(t),
          save: async (g) => props.onSaved(await workingApi.saveModel(model.name, { ...model.info, manualGroup: g }, model.revision)),
        });
      } catch (e) { out.processError = errorText(e); }
    }
    if (file) {
      setStatus("JIG/MC List: the AI is matching the workbook (1-3 minutes)…");
      try {
        out.jig = await runJig({
          match: () => processListApi.aiMatch(model.name, file),
          current: () => processListApi.get(model.name),
          confirm: (t) => window.confirm(t),
          save: async (rows, revision) => { await processListApi.save(model.name, rows, revision); },
        });
      } catch (e) { out.jigError = errorText(e); }
    }
    setResult(out); setStatus(""); setBusy(false);
  };

  return (
    <div className="wk-docs">
      <section className="card pi-sec">
        <div className="pi-sec-h"><h3>Work Instructions</h3></div>
        <div className="pi-sec-b"><WorkInstructionTab modelName={model.name} onError={props.onError} /></div>
      </section>
      <section className="card pi-sec">
        <div className="pi-sec-h"><h3>JIG / Machine / Tester list</h3></div>
        <div className="pi-sec-b">
          <input ref={pick} type="file" accept=".xlsx,.xlsm" hidden aria-label="JIG / Machine / Tester workbook"
            onChange={(e) => { const f = e.target.files?.[0]; setFile(isWorkbook(f) ? f! : null); if (f && !isWorkbook(f)) setStatus("Choose an .xlsx or .xlsm file."); }} />
          <button type="button" className="btn sm" disabled={busy} onClick={() => pick.current?.click()}>Choose workbook</button>
          <span className="small muted wk-docs-file">{file ? file.name : "No file chosen"}</span>
        </div>
      </section>
      <button type="button" className="btn primary wk-docs-ai" disabled={busy} onClick={() => void run()}>✦ AI: make Process and JIG/MC List</button>
      {status && <div className="small wk-docs-status" role="status">{status}</div>}
      {result && (
        <div className="small wk-docs-result">
          {result.process && (
            <div>
              <b>{processSummary(result.process)}</b>{result.process.saved ? " · saved" : " · not saved"}
              {result.process.note && <span className="muted"> · {result.process.note}</span>}
              {" "}<button type="button" className="btn sm" onClick={() => props.onOpen("process")}>Open Process</button>
            </div>
          )}
          {result.processError && <div className="pi-dirty">Process: {result.processError}</div>}
          {result.jig && (
            <div>
              <b>{jigSummary(result.jig)}</b>{result.jig.saved ? " · saved" : " · not saved"}
              {result.jig.unmatched > 0 && <span className="pi-dirty"> · {result.jig.unmatched} unmatched</span>}
              {result.jig.note && <span className="muted"> · {result.jig.note}</span>}
              {" "}<button type="button" className="btn sm" onClick={() => props.onOpen("jigmc")}>Open JIG/MC List</button>
            </div>
          )}
          {result.jigError && <div className="pi-dirty">JIG/MC List: {result.jigError}</div>}
        </div>
      )}
    </div>
  );
}
