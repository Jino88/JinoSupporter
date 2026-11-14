import { useRef, useState } from "react";
import WorkInstructionTab from "./WorkInstructionTab";
import { workingApi } from "./api";
import { loadInstructions, uploadInstructions } from "./workInstruction";
import { readManual } from "./manual";
import { processListApi } from "./processList";
import {
  KIND_TEXT, NOTHING_TO_DO, docKind, docLinks, errorText, isWorkbook, jigSummary, manualAiFromInstructions, processSummary, runJig, runProcess, sortDocs,
  type DocKind, type DocsResult, type FoundKind,
} from "./docsAi";
import type { WorkingModel } from "./types";
import "./modelDocs.css";

/** The relation the documents make: each process of the line with its materials (Process) and its machine and JIGs (JIG/MC List). */
export function DocLinks(props: { result: DocsResult }) {
  const links = docLinks(props.result.process?.group, props.result.jig?.rows);
  if (!links.length) return null;
  const both = links.filter((l) => l.inProcess && l.inJig).length;
  return (
    <div className="wk-docs-links">
      <div className="small"><b>How the documents relate</b> · {both} of {links.length} processes are in both Process and JIG/MC List</div>
      <table className="wk-it">
        <thead><tr><th>No.</th><th className="l">Process</th><th className="l">Materials (Process)</th><th className="l">Machine</th><th className="l">JIG</th></tr></thead>
        <tbody>
          {links.map((l, i) => (
            <tr key={i} className={l.inProcess && l.inJig ? "" : "wk-docs-half"}>
              <td>{l.no}</td>
              <td className="l"><b>{l.process}</b>{!l.inJig && <span className="muted"> · not in JIG/MC List</span>}{!l.inProcess && <span className="muted"> · not in Process</span>}</td>
              <td className="l">{l.materials.join(", ")}</td>
              <td className="l wk-docs-pre">{l.machine}</td>
              <td className="l wk-docs-pre">{l.jigs.join("\n")}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/**
 * Model Info › Documents: one AI button takes every document of the model at once (Work Instructions, JIG / Machine /
 * Tester list), finds what each is, keeps the Work Instructions, makes the Process table from them and the JIG/MC List
 * in their line order, and shows how the documents relate. The Work Instructions are shown below.
 */
export default function ModelDocs(props: {
  model: WorkingModel; onSaved: (m: WorkingModel) => void; onOpen: (k: "process" | "jigmc") => void; onError: (e: unknown) => void;
  /** The Process table or the JIG/MC List was saved (their tabs read it again). */
  onDocsSaved?: () => void;
}) {
  const { model } = props;
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);
  const [found, setFound] = useState<{ name: string; kind: DocKind }[]>([]);
  const [skipped, setSkipped] = useState<string[]>([]);
  const [result, setResult] = useState<DocsResult | null>(null);
  const [fresh, setFresh] = useState(0);
  const pick = useRef<HTMLInputElement>(null);

  const run = async (files: File[]) => {
    if (busy) return;
    setResult(null); setFound([]); setSkipped([]);
    const books = files.filter((f) => isWorkbook(f));
    const others = files.filter((f) => !isWorkbook(f)).map((f) => `${f.name} (not an .xlsx / .xlsm workbook)`);
    setBusy(true);
    try {
      // what each document is
      setStatus("Reading what each document is…");
      const kinds: ({ file: File } & FoundKind)[] = [];
      for (const f of books) kinds.push({ file: f, ...(await docKind(model.name, f)) });
      setFound(kinds.map((k) => ({ name: k.file.name, kind: k.kind })));
      const docs = sortDocs(kinds);
      setSkipped([...others, ...docs.skipped]);

      // the Work Instructions first: the Process table and the line order of the JIG/MC List come from them
      if (docs.instructions) {
        setStatus(`Keeping the Work Instructions ${docs.instructions.name}…`);
        await uploadInstructions(model.name, docs.instructions);
        setFresh((n) => n + 1);
      }
      const hasInstructions = !!docs.instructions || (await loadInstructions(model.name)) !== null;
      if (!hasInstructions && !docs.jig) { setStatus(NOTHING_TO_DO); return; }

      const out: DocsResult = {};
      if (hasInstructions) {
        setStatus("Process: the AI is reading the Work Instructions…");
        try {
          out.process = await runProcess({
            hasInstructions,
            currentGroup: readManual(model.info.manualGroup),
            ask: (g) => manualAiFromInstructions(model.name, g),
            confirm: (t) => window.confirm(t),
            save: async (g) => { props.onSaved(await workingApi.saveModel(model.name, { ...model.info, manualGroup: g }, model.revision)); props.onDocsSaved?.(); },
          });
        } catch (e) { out.processError = errorText(e); }
      }
      if (docs.jig) {
        const jig = docs.jig;
        setStatus("JIG/MC List: the AI is putting the machines and JIGs on the processes (1-3 minutes)…");
        try {
          out.jig = await runJig({
            match: () => processListApi.aiMatch(model.name, jig),
            current: () => processListApi.get(model.name),
            confirm: (t) => window.confirm(t),
            save: async (rows, revision) => { await processListApi.save(model.name, rows, revision); props.onDocsSaved?.(); },
          });
        } catch (e) { out.jigError = errorText(e); }
      }
      setResult(out); setStatus("");
    } catch (e) {
      setStatus(errorText(e));
    } finally { setBusy(false); }
  };

  return (
    <div className="wk-docs">
      <section className="card pi-sec">
        <div className="pi-sec-h"><h3>Documents</h3></div>
        <div className="pi-sec-b">
          <input ref={pick} type="file" accept=".xlsx,.xlsm" multiple hidden aria-label="Documents of the model"
            onChange={(e) => { const list = [...(e.target.files ?? [])]; e.target.value = ""; if (list.length) void run(list); }} />
          <div className="wk-docs-bar">
            <button type="button" className="btn primary wk-docs-ai" disabled={busy} onClick={() => pick.current?.click()}>
              {busy ? "AI is working…" : "✦ AI: read the documents"}
            </button>
            <span className="small muted">
              Pick every document of the model at once (Work Instructions, JIG / Machine / Tester list). AI finds what each is,
              makes the Process table and the JIG/MC List from them, and shows how they relate.
            </span>
          </div>
          {status && <div className="small wk-docs-status" role="status">{status}</div>}
          {found.length > 0 && (
            <ul className="small wk-docs-found">
              {found.map((f, i) => <li key={i}><b>{f.name}</b> → {KIND_TEXT[f.kind]}</li>)}
            </ul>
          )}
          {skipped.length > 0 && <div className="small muted">Not used: {skipped.join(", ")}</div>}
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
                  {result.jig.unmatched > 0 && <span className="muted"> · {result.jig.unmatched} items not put on a process</span>}
                  {result.jig.note && <span className="muted"> · {result.jig.note}</span>}
                  {" "}<button type="button" className="btn sm" onClick={() => props.onOpen("jigmc")}>Open JIG/MC List</button>
                </div>
              )}
              {result.jigError && <div className="pi-dirty">JIG/MC List: {result.jigError}</div>}
              <DocLinks result={result} />
            </div>
          )}
        </div>
      </section>
      {/* the Work Instructions kept, to read (read again after the AI keeps a new file) */}
      <WorkInstructionTab key={fresh} modelName={model.name} onError={props.onError} />
    </div>
  );
}
