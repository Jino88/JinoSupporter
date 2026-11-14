import { Fragment, useEffect, useRef, useState } from "react";
import { downloadBlob } from "../../toolsApi";
import { useTr } from "./i18n";
import {
  deleteInstructions, equipmentText, instructionFileUrl, loadInstructions, sizeText, timeText, uploadInstructions, whenText,
  type InstructionDoc, type InstructionProcess,
} from "./workInstruction";
import "./workInstruction.css";

/** The detail of one process page: parts, equipment & tools, steps and the key control items. */
export function ProcessDetail({ p }: { p: InstructionProcess }) {
  const tr = useTr();
  return (
    <div className="wi-detail">
      <div>
        <h4>{tr("Parts")}</h4>
        {p.parts.length ? (
          <table className="wk-it wi-t">
            <thead><tr><th>NO</th><th className="l">{tr("Part name")}</th><th className="l">{tr("Spec")}</th><th>{tr("Qty")}</th></tr></thead>
            <tbody>{p.parts.map((x, i) => <tr key={i}><td>{x.no}</td><td className="l">{x.name}</td><td className="l">{x.spec}</td><td>{x.qty}</td></tr>)}</tbody>
          </table>
        ) : <div className="muted">{tr("None")}</div>}
      </div>
      <div>
        <h4>{tr("Equipment & tools")}</h4>
        {p.tools.length ? (
          <table className="wk-it wi-t">
            <thead><tr><th>NO</th><th className="l">{tr("Name")}</th><th className="l">{tr("Note")}</th></tr></thead>
            <tbody>{p.tools.map((x, i) => <tr key={i}><td>{x.no}</td><td className="l">{x.name}</td><td className="l">{x.note}</td></tr>)}</tbody>
          </table>
        ) : <div className="muted">{tr("None")}</div>}
      </div>
      <div>
        <h4>{tr("Steps")}</h4>
        {p.steps.length
          ? <ol>{p.steps.map((s, i) => <li key={i} value={s.no}>{s.text}</li>)}</ol>
          : <div className="muted">{tr("None")}</div>}
      </div>
      <div>
        <h4>{tr("Key control items")}</h4>
        {p.keyPoints.length
          ? <ul>{p.keyPoints.map((k, i) => <li key={i}>{k}</li>)}</ul>
          : <div className="muted">{tr("None")}</div>}
      </div>
    </div>
  );
}

/** What the tab shows for a model: the upload prompt, or the file, its revision history and the process table with the open page. */
export function InstructionBody(props: {
  doc: InstructionDoc | null;
  busy: boolean;
  open: string | null;
  onOpen: (sheet: string | null) => void;
  onPick: () => void;
  onDownload: () => void;
  onDelete: () => void;
}) {
  const tr = useTr();
  const { doc } = props;
  if (!doc)
    return (
      <div className="wi-empty">
        <span>{tr("No work instruction workbook is uploaded for this model yet. Upload the Excel file (.xlsx); each sheet with a process name becomes one process page.")}</span>
        <button className="btn sm primary" disabled={props.busy} onClick={props.onPick}>{props.busy ? tr("Reading…") : tr("Upload workbook")}</button>
      </div>
    );
  return (
    <>
      <div className="wi-file">
        <b>{doc.fileName}</b>
        <span className="muted">{whenText(doc.uploadedAt)} · {doc.uploadedBy} · {sizeText(doc.size)}</span>
        <span className="btns">
          <button className="btn sm" disabled={props.busy} onClick={props.onPick}>{props.busy ? tr("Reading…") : tr("Replace")}</button>
          <button className="btn sm" disabled={props.busy} onClick={props.onDownload}>{tr("Download")}</button>
          <button className="btn sm" disabled={props.busy} onClick={props.onDelete}>{tr("Delete")}</button>
        </span>
      </div>
      {doc.revisions.length > 0 && (
        <>
          <div className="wi-sub">{tr("Revision history")}</div>
          <div className="pi-wrap">
            <table className="wk-it wi-t">
              <thead><tr><th>{tr("Date")}</th><th>{tr("Page")}</th><th>{tr("Rev.")}</th><th className="l">{tr("Details")}</th><th className="l">{tr("Author")}</th></tr></thead>
              <tbody>
                {doc.revisions.map((r, i) => (
                  <tr key={i}><td>{r.date}</td><td>{r.page}</td><td>{r.rev}</td><td className="l pre">{r.content}</td><td className="l">{r.author}</td></tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
      <div className="wi-sub">{tr("Processes")} ({doc.processes.length})</div>
      <div className="pi-wrap">
        <table className="wk-it wi-t">
          <thead>
            <tr><th>No</th><th>{tr("Code")}</th><th className="l">{tr("Process")}</th><th>{tr("Standard time")}</th><th className="l">{tr("Equipment")}</th><th>CTQ</th></tr>
          </thead>
          <tbody>
            {doc.processes.map((p, i) => (
              <Fragment key={p.sheet + i}>
                <tr className={"wi-row" + (props.open === p.sheet ? " on" : "")} onClick={() => props.onOpen(props.open === p.sheet ? null : p.sheet)}>
                  <td>{i + 1}</td><td>{p.code}</td><td className="l">{p.name}</td><td>{timeText(p.standardTime)}</td>
                  <td className="l">{equipmentText(p)}</td><td>{p.ctq ? <span className="wi-ctq">CTQ</span> : ""}</td>
                </tr>
                {props.open === p.sheet && <tr><td colSpan={6}><ProcessDetail p={p} /></td></tr>}
              </Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}

/** Work Instructions (작업지도서) of a model: the pages of its work instruction workbook, one per process. */
export default function WorkInstructionTab(props: { modelName: string; onError: (e: unknown) => void }) {
  const { modelName, onError } = props;
  const tr = useTr();
  const [doc, setDoc] = useState<InstructionDoc | null | undefined>(undefined);
  const [busy, setBusy] = useState(false);
  const [open, setOpen] = useState<string | null>(null);
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    let live = true;
    loadInstructions(modelName).then((d) => { if (live) setDoc(d); }, (e) => { if (live) { setDoc(null); onError(e); } });
    return () => { live = false; };
  }, [modelName]); // eslint-disable-line react-hooks/exhaustive-deps

  const send = async (file: File | undefined) => {
    if (!file) return;
    setBusy(true);
    try { setDoc(await uploadInstructions(modelName, file)); setOpen(null); }
    catch (e) { onError(e); }
    finally { setBusy(false); }
  };
  const download = async () => {
    try {
      const response = await fetch(instructionFileUrl(modelName), { credentials: "same-origin" });
      if (!response.ok) throw new Error(`Request failed (HTTP ${response.status})`);
      downloadBlob(await response.blob(), doc?.fileName ?? "work-instructions.xlsx");
    } catch (e) { onError(e); }
  };
  const remove = async () => {
    if (!window.confirm(tr("Delete the work instruction workbook of this model?"))) return;
    setBusy(true);
    try { await deleteInstructions(modelName); setDoc(null); setOpen(null); }
    catch (e) { onError(e); }
    finally { setBusy(false); }
  };

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h"><h2>{tr("Work Instructions")}</h2></div>
      <div className="pi-sec-b">
        <input ref={input} type="file" accept=".xlsx,.xlsm" hidden
          onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ""; void send(f); }} />
        {doc === undefined
          ? <div className="muted">{tr("Loading…")}</div>
          : <InstructionBody doc={doc} busy={busy} open={open} onOpen={setOpen} onPick={() => input.current?.click()} onDownload={() => void download()} onDelete={() => void remove()} />}
      </div>
    </section>
  );
}
