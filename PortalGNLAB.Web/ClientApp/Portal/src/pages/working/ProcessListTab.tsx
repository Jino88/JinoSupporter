import { useEffect, useRef, useState } from "react";
import { useTr } from "./i18n";
import { BASIS_TEXT, blankRow, lineCount, nextNo, normalizeList, normalizeRow, processListApi, rowsProblem } from "./processList";
import type { ProcessList, ProcessMatch, ProcessRow } from "./processList";
import "./processList.css";

const stamp = (iso: string | null) => (iso ? iso.slice(0, 16).replace("T", " ") : "");
// Machine, JIG, Ver., Qty and Tester hold one item per line; line i of JIG, Ver. and Qty is the same JIG.
const TEXT_COLS = [["machine", "Machine"], ["jig", "JIG"], ["jigVer", "Ver."], ["jigQty", "Qty"], ["tester", "Tester"], ["material", "Material"], ["note", "Note"]] as const;
const NARROW = new Set<string>(["jigVer", "jigQty"]);

/** One row being edited: every cell an input. */
function EditRow(props: { row: ProcessRow; onChange: (r: ProcessRow) => void; onRemove: () => void }) {
  const r = props.row;
  const set = <K extends keyof ProcessRow>(k: K, v: ProcessRow[K]) => props.onChange({ ...r, [k]: v });
  return (
    <tr className="pl-edit">
      <td><input className="inp sm pl-no" aria-label="No." value={r.no} onChange={(e) => set("no", e.target.value)} /></td>
      <td className="l"><input className="inp sm pl-proc" aria-label="Process" value={r.process} onChange={(e) => set("process", e.target.value)} /></td>
      {TEXT_COLS.map(([k, label]) => (
        <td key={k} className="l"><textarea className={`inp sm pl-text${NARROW.has(k) ? " pl-narrow" : ""}`} rows={lineCount(r[k])} aria-label={label} value={r[k]} onChange={(e) => set(k, e.target.value)} /></td>
      ))}
      <td className="act"><button className="btn sm" title="Remove the row" onClick={props.onRemove}>✕</button></td>
    </tr>
  );
}

/** What the AI match says about its rows: where the order came from, the model, its note and what it could not place. */
function MatchNotice(props: { match: ProcessMatch }) {
  const tr = useTr();
  const m = props.match;
  return (
    <div className="pl-ai" role="note">
      <div><b>✦ {tr(BASIS_TEXT[m.basis] ?? BASIS_TEXT.workbook)}</b>{m.model && <span className="muted small"> · {m.model}</span>}</div>
      {m.note && <div className="pl-ai-note">{m.note}</div>}
      {m.unmatched.length > 0 && (
        <details className="pl-ai-un">
          <summary>{tr("Not placed in a process")} ({m.unmatched.length})</summary>
          <ul>{m.unmatched.map((u, i) => <li key={i}>{u}</li>)}</ul>
        </details>
      )}
    </div>
  );
}

/** Process List (공정편성 LIST) of a model: one row per process with its machines, jigs (version, quantity) and testers. */
export default function ProcessListTab(props: {
  modelName: string;
  onError: (e: unknown) => void;
  /** List already read (tests); else read here. */
  initial?: ProcessList;
  /** Start in edit mode (tests). */
  editing?: boolean;
  /** AI match result already in the draft (tests). */
  match?: ProcessMatch;
  /** Start while the AI is matching (tests). */
  matching?: boolean;
}) {
  const tr = useTr();
  const [list, setList] = useState<ProcessList | undefined>(props.initial && normalizeList(props.initial));
  const [draft, setDraft] = useState<ProcessRow[] | null>(
    props.match ? props.match.rows.map(normalizeRow) : props.editing && props.initial ? props.initial.rows.map(normalizeRow) : null);
  const [problem, setProblem] = useState<string | null>(null);
  const [busy, setBusy] = useState(props.matching ?? false);
  const [matching, setMatching] = useState(props.matching ?? false);
  const [match, setMatch] = useState<ProcessMatch | null>(props.match ?? null);
  const input = useRef<HTMLInputElement>(null);
  const aiInput = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (props.initial !== undefined) return;
    let live = true;
    processListApi.get(props.modelName).then((l) => { if (live) setList(normalizeList(l)); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  if (!list) return <section className="card pi-sec"><div className="pi-sec-b muted">{tr("Loading…")}</div></section>;

  const run = async (work: () => Promise<ProcessList>) => {
    setBusy(true);
    try { setList(normalizeList(await work())); setDraft(null); setMatch(null); setProblem(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
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

  const aiMatch = async (file: File | undefined) => {
    if (aiInput.current) aiInput.current.value = "";
    if (!file) return;
    if (list.rows.length > 0 && !window.confirm(tr("Replace the rows of the Process List with the ones the AI matches from the workbook? You can review them before saving."))) return;
    setBusy(true); setMatching(true);
    try {
      const m = await processListApi.aiMatch(props.modelName, file);
      if (m.rows.length > 0) { setDraft(m.rows.map(normalizeRow)); setMatch(m); setProblem(null); }
      else setProblem(m.note || "The AI found no processes to match in the workbook.");
    } catch (e) { props.onError(e); } finally { setBusy(false); setMatching(false); }
  };

  const rows = draft ?? list.rows;
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
              <button className="btn sm" disabled={busy} onClick={() => { setDraft(null); setMatch(null); setProblem(null); }}>{tr("Cancel")}</button>
              <button className="btn sm" disabled={busy} onClick={() => setDraft([...draft, blankRow(nextNo(draft))])}>{tr("+ Row")}</button>
              {problem && <span className="pl-err">{tr(problem)}</span>}
            </>
          ) : (
            <>
              <button className="btn sm" disabled={busy} onClick={() => setDraft(list.rows.length ? list.rows : [blankRow("01")])}>{tr("Edit")}</button>
              <button className="btn sm" disabled={busy} title={tr("Read the Machine List, JIG LIST and Tester List sheets of the workbook")}
                onClick={() => input.current?.click()}>{busy && !matching ? tr("Reading the workbook…") : tr("Import JIG / Machine list")}</button>
              <input ref={input} type="file" accept=".xlsx,.xlsm" hidden onChange={(e) => importFile(e.target.files?.[0])} />
              <button className="btn sm" disabled={busy} aria-busy={matching}
                title={tr("The AI matches the machines, JIGs and testers of the workbook to the processes, in process order")}
                onClick={() => aiInput.current?.click()}>{matching ? tr("AI is matching…") : tr("✦ AI match")}</button>
              <input ref={aiInput} type="file" accept=".xlsx,.xlsm" hidden aria-label="AI match workbook" onChange={(e) => aiMatch(e.target.files?.[0])} />
              {problem && <span className="pl-err">{tr(problem)}</span>}
            </>
          )}
          <span className="pl-sum muted small">{tr("Processes")} <b>{rows.length}</b></span>
        </div>
        {draft && match && <MatchNotice match={match} />}
        {rows.length === 0 && !draft ? (
          <div className="muted">{tr("No processes yet. Import the JIG / Machine / Tester list workbook of the model, or Edit to type them in.")}</div>
        ) : (
          <div className="pl-wrap">
            <table className="wk-it pl-t">
              <thead>
                <tr>
                  <th>No.</th><th className="l">{tr("Process")}</th>
                  {TEXT_COLS.map(([k, label]) => <th key={k} className={NARROW.has(k) ? "pl-narrow" : "l"}>{tr(label)}</th>)}
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
                      {TEXT_COLS.map(([k]) => <td key={k} className={`${NARROW.has(k) ? "" : "l "}pl-cell${NARROW.has(k) ? " pl-narrow" : ""}`}>{r[k]}</td>)}
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
