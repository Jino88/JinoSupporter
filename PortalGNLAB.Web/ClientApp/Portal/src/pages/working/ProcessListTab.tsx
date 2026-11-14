import { useEffect, useRef, useState } from "react";
import { useTr } from "./i18n";
import { blankRow, cellNumber, nextNo, processListApi, rowsProblem, totals } from "./processList";
import type { ProcessList, ProcessRow } from "./processList";
import "./processList.css";

const stamp = (iso: string | null) => (iso ? iso.slice(0, 16).replace("T", " ") : "");
const TEXT_COLS = [["machine", "Machine"], ["jig", "JIG"], ["tester", "Tester"], ["material", "Material"], ["note", "Note"]] as const;

/** One row being edited: every cell an input. */
function EditRow(props: { row: ProcessRow; onChange: (r: ProcessRow) => void; onRemove: () => void }) {
  const r = props.row;
  const set = <K extends keyof ProcessRow>(k: K, v: ProcessRow[K]) => props.onChange({ ...r, [k]: v });
  const num = (k: "workers" | "stdTime", whole: boolean, label: string) => (
    <td><input className="inp sm pl-num" type="number" min={0} step={whole ? 1 : "any"} aria-label={label} value={r[k] ?? ""}
      onChange={(e) => { const n = cellNumber(e.target.value, whole); if (n !== undefined) set(k, n); }} /></td>
  );
  return (
    <tr className="pl-edit">
      <td><input className="inp sm pl-no" aria-label="No." value={r.no} onChange={(e) => set("no", e.target.value)} /></td>
      <td className="l"><input className="inp sm pl-proc" aria-label="Process" value={r.process} onChange={(e) => set("process", e.target.value)} /></td>
      {num("workers", true, "Workers")}
      {num("stdTime", false, "Standard time")}
      {TEXT_COLS.map(([k, label]) => (
        <td key={k} className="l"><textarea className="inp sm pl-text" rows={1} aria-label={label} value={r[k]} onChange={(e) => set(k, e.target.value)} /></td>
      ))}
      <td className="act"><button className="btn sm" title="Remove the row" onClick={props.onRemove}>✕</button></td>
    </tr>
  );
}

/** Process List (공정편성 LIST) of a model: one row per process with its people, standard time, machine and jigs. */
export default function ProcessListTab(props: {
  modelName: string;
  onError: (e: unknown) => void;
  /** List already read (tests); else read here. */
  initial?: ProcessList;
  /** Start in edit mode (tests). */
  editing?: boolean;
}) {
  const tr = useTr();
  const [list, setList] = useState<ProcessList | undefined>(props.initial);
  const [draft, setDraft] = useState<ProcessRow[] | null>(props.editing && props.initial ? props.initial.rows : null);
  const [problem, setProblem] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (props.initial !== undefined) return;
    let live = true;
    processListApi.get(props.modelName).then((l) => { if (live) setList(l); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!list) return <section className="card pi-sec"><div className="pi-sec-b muted">{tr("Loading…")}</div></section>;

  const run = async (work: () => Promise<ProcessList>) => {
    setBusy(true);
    try { setList(await work()); setDraft(null); setProblem(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const save = () => {
    if (!draft) return;
    const p = rowsProblem(draft);
    setProblem(p);
    if (!p) run(() => processListApi.save(props.modelName, draft, list.revision));
  };
  const importFile = (file: File | undefined) => {
    if (input.current) input.current.value = "";
    if (!file) return;
    if (list.rows.length > 0 && !window.confirm(tr("Replace the rows of the Process List with the ones read from the workbook?"))) return;
    run(() => processListApi.import(props.modelName, file, list.revision));
  };

  const rows = draft ?? list.rows;
  const t = totals(rows);
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Process List</h2>
        {list.updatedAt && (
          <span className="muted small">
            {list.source && <><b>{list.source}</b> · </>}{stamp(list.updatedAt)} · {list.updatedBy}
          </span>
        )}
      </div>
      <div className="pi-sec-b">
        <div className="pl-bar">
          {draft ? (
            <>
              <button className="btn sm primary" disabled={busy} onClick={save}>{tr("Save")}</button>
              <button className="btn sm" disabled={busy} onClick={() => { setDraft(null); setProblem(null); }}>{tr("Cancel")}</button>
              <button className="btn sm" disabled={busy} onClick={() => setDraft([...draft, blankRow(nextNo(draft))])}>{tr("+ Row")}</button>
              {problem && <span className="pl-err">{tr(problem)}</span>}
            </>
          ) : (
            <>
              <button className="btn sm" disabled={busy} onClick={() => setDraft(list.rows.length ? list.rows : [blankRow("01")])}>{tr("Edit")}</button>
              <button className="btn sm" disabled={busy} title={tr("Read the Machine List, JIG LIST and Tester List sheets of the workbook")}
                onClick={() => input.current?.click()}>{busy ? tr("Reading the workbook…") : tr("Import JIG / Machine list")}</button>
              <input ref={input} type="file" accept=".xlsx,.xlsm" hidden onChange={(e) => importFile(e.target.files?.[0])} />
            </>
          )}
          <span className="pl-sum muted small">
            {tr("Processes")} <b>{t.processes}</b> · {tr("Workers")} <b>{t.workers}</b>
            {t.slowest && <> · {tr("Slowest")} <b>{t.slowest.stdTime}s</b> ({t.slowest.no} {t.slowest.process})</>}
          </span>
        </div>
        {rows.length === 0 && !draft ? (
          <div className="muted">{tr("No processes yet. Import the JIG / Machine / Tester list workbook of the model, or Edit to type them in.")}</div>
        ) : (
          <div className="pl-wrap">
            <table className="wk-it pl-t">
              <thead>
                <tr>
                  <th>No.</th><th className="l">{tr("Process")}</th><th>{tr("Workers")}</th><th>{tr("Std. time (s)")}</th>
                  {TEXT_COLS.map(([k, label]) => <th key={k} className="l">{tr(label)}</th>)}
                  {draft && <th></th>}
                </tr>
              </thead>
              <tbody>
                {draft
                  ? draft.map((r, i) => (
                    <EditRow key={i} row={r} onChange={(n) => setDraft(draft.map((x, j) => (j === i ? n : x)))}
                      onRemove={() => setDraft(draft.filter((_, j) => j !== i))} />
                  ))
                  : list.rows.map((r) => (
                    <tr key={r.id}>
                      <td className="pl-no">{r.no}</td>
                      <td className="l"><b>{r.process}</b></td>
                      <td>{r.workers ?? ""}</td>
                      <td>{r.stdTime ?? ""}</td>
                      {TEXT_COLS.map(([k]) => <td key={k} className="l pl-cell">{r[k]}</td>)}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </section>
  );
}
