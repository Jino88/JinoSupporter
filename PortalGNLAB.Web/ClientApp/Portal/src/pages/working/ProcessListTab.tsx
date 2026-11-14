import { useEffect, useRef, useState } from "react";
import { useTr } from "./i18n";
import { BASIS_TEXT, addJig, blankRow, jigTotals, lineCount, nextNo, normalizeList, normalizeRow, processListApi, removeJig, rowsProblem, setJig } from "./processList";
import type { Jig, ProcessList, ProcessMatch, ProcessRow } from "./processList";
import "./processList.css";

const stamp = (iso: string | null) => (iso ? iso.slice(0, 16).replace("T", " ") : "");
// Machine and Tester hold one item per line; every JIG is a sub-row of its process with its own Drawing No., Ver. and Qty.
type TextKey = "machine" | "tester" | "material" | "note";
const HEAD_COLS = [["machine", "Machine"]] as const;
const TAIL_COLS = [["tester", "Tester"], ["material", "Material"], ["note", "Note"]] as const;
/** Cells of one item per line: a line is never wrapped. */
const LINES = new Set<string>(["machine", "tester"]);

/** A text cell of the process: plain text to read, a box as big as its text to edit. */
function TextCell(props: { k: TextKey; label: string; row: ProcessRow; span: number; edit: boolean; onChange: (r: ProcessRow) => void }) {
  const { k, row: r } = props;
  const lines = LINES.has(k);
  return (
    <td rowSpan={props.span} className={`l pl-cell${lines ? " pl-lines" : ""}`}>
      {props.edit
        ? <textarea className="inp sm pl-text" wrap={lines ? "off" : undefined} rows={lineCount(r[k])} aria-label={props.label}
          value={r[k]} onChange={(e) => props.onChange({ ...r, [k]: e.target.value })} />
        : r[k]}
    </td>
  );
}

/** One process: a row per JIG (one row when it has none); the cells of the process span them all. */
function ProcessRows(props: { row: ProcessRow; edit: boolean; onChange: (r: ProcessRow) => void; onRemove: () => void }) {
  const { row: r, edit } = props;
  const tr = useTr();
  const span = Math.max(1, r.jigs.length);
  const set = <K extends keyof ProcessRow>(k: K, v: ProcessRow[K]) => props.onChange({ ...r, [k]: v });
  const items: (Jig | null)[] = r.jigs.length ? r.jigs : [null];
  return (
    <>
      {items.map((j, i) => {
        const jigInput = (k: keyof Jig, label: string, extra = "") => j && (
          <input className={`inp sm pl-jig${extra}`} aria-label={label} value={j[k]} onChange={(e) => props.onChange(setJig(r, i, { [k]: e.target.value }))} />
        );
        return (
          <tr key={i} className={`${i === 0 ? "pl-first" : "pl-sub"}${edit ? " pl-edit" : ""}`}>
            {i === 0 && (
              <>
                <td rowSpan={span} className="pl-no">{edit ? <input className="inp sm pl-no" aria-label="No." value={r.no} onChange={(e) => set("no", e.target.value)} /> : r.no}</td>
                <td rowSpan={span} className="l">{edit ? <input className="inp sm pl-proc" aria-label="Process" value={r.process} onChange={(e) => set("process", e.target.value)} /> : <b>{r.process}</b>}</td>
                {HEAD_COLS.map(([k, label]) => <TextCell key={k} k={k} label={label} row={r} span={span} edit={edit} onChange={props.onChange} />)}
              </>
            )}
            <td className="l pl-jigcell">{edit ? jigInput("name", "JIG") : j?.name}</td>
            <td className="l pl-jigcell">{edit ? jigInput("drawingNo", "Drawing No") : j?.drawingNo}</td>
            <td className="pl-narrow pl-jigcell">{edit ? jigInput("ver", "Ver.", " pl-narrow") : j?.ver}</td>
            <td className="pl-narrow pl-jigcell">{edit ? jigInput("qty", "Qty", " pl-narrow") : j?.qty}</td>
            {edit && (
              <td className="act">{j && <button className="btn sm" title={tr("Remove the JIG")} aria-label="Remove the JIG" onClick={() => props.onChange(removeJig(r, i))}>✕</button>}</td>
            )}
            {i === 0 && (
              <>
                {TAIL_COLS.map(([k, label]) => <TextCell key={k} k={k} label={label} row={r} span={span} edit={edit} onChange={props.onChange} />)}
                {edit && (
                  <td rowSpan={span} className="act">
                    <button className="btn sm" onClick={() => props.onChange(addJig(r))}>{tr("+ JIG")}</button>{" "}
                    <button className="btn sm" title={tr("Remove the row")} aria-label="Remove the row" onClick={props.onRemove}>✕</button>
                  </td>
                )}
              </>
            )}
          </tr>
        );
      })}
    </>
  );
}

/** All JIGs of the list as one, always shown: the count per kind of JIG and the total, above the processes. */
function JigSummary(props: { rows: ProcessRow[] }) {
  const tr = useTr();
  const t = jigTotals(props.rows);
  return (
    <div className="pl-jigsum">
      <div className="pl-jigsum-h">{tr("JIG total")}: <b>{t.kinds}</b>{tr(" kinds")} · <b>{t.total}</b>{tr(" pcs")}</div>
      {t.items.length > 0 && (
        <div className="pl-jigtot">
          {t.items.map((j, i) => (
            <div key={i} className="pl-jigtot-i">
              <span className="pl-jigtot-n">{j.name}</span>
              <span className="pl-jigtot-d">{j.drawingNo}</span>
              <span className="pl-jigtot-v">{j.ver}</span>
              <b className="pl-jigtot-q">{j.qtyNum}</b>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

/**
 * What the AI match says about its rows, as information (not an error): the rows are not saved yet, where the order
 * came from, the model, its note and the items it could not put on a process with the reason for each.
 */
function MatchNotice(props: { match: ProcessMatch }) {
  const tr = useTr();
  const m = props.match;
  return (
    <div className="pl-ai" role="note">
      <div>
        <b>✦ {tr("AI match result (not saved yet)")}</b>
        <span className="muted small"> · {tr(BASIS_TEXT[m.basis] ?? BASIS_TEXT.workbook)}{m.model && <> · {m.model}</>}</span>
      </div>
      {m.note && <div className="pl-ai-note"><span className="muted">{tr("AI note")}:</span> {m.note}</div>}
      {m.unmatched.length > 0 && (
        <details className="pl-ai-un">
          <summary>{tr("Items not put on a process")} ({m.unmatched.length})</summary>
          <div className="muted small">
            {tr("They are for a process that is not in this line's process list, or have no process number. Add them with Edit if needed.")}
          </div>
          <ul>{m.unmatched.map((u, i) => <li key={i}>{u}</li>)}</ul>
        </details>
      )}
    </div>
  );
}

/** Process List (공정편성 LIST) of a model: one row per process with its machines, jigs (drawing no., version, quantity) and testers. */
export default function ProcessListTab(props: {
  modelName: string;
  onError: (e: unknown) => void;
  /** List already read (tests); else read here. */
  initial?: ProcessList;
  /** Start in edit mode with the rows of `initial` as the draft (tests). */
  editing?: boolean;
  /** AI match result already in the draft, read as text until Edit (tests). */
  match?: ProcessMatch;
  /** Start while the AI is matching (tests). */
  matching?: boolean;
}) {
  const tr = useTr();
  const [list, setList] = useState<ProcessList | undefined>(props.initial && normalizeList(props.initial));
  const [draft, setDraft] = useState<ProcessRow[] | null>(
    props.match ? props.match.rows.map(normalizeRow) : props.editing && props.initial ? normalizeList(props.initial).rows : null);
  // a draft is read as text until Edit is pressed (an AI match result starts as such a draft)
  const [editMode, setEditMode] = useState(Boolean(props.editing && props.initial && !props.match));
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

  const leave = () => { setDraft(null); setEditMode(false); setMatch(null); setProblem(null); };
  const run = async (work: () => Promise<ProcessList>) => {
    setBusy(true);
    try { setList(normalizeList(await work())); leave(); } catch (e) { props.onError(e); } finally { setBusy(false); }
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
      if (m.rows.length > 0) { setDraft(m.rows.map(normalizeRow)); setEditMode(false); setMatch(m); setProblem(null); }
      else setProblem(m.note || "The AI found no processes to match in the workbook.");
    } catch (e) { props.onError(e); } finally { setBusy(false); setMatching(false); }
  };

  const rows = draft ?? list.rows;
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>JIG/MC List</h2>
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
              <button className="btn sm" disabled={busy} onClick={leave}>{tr("Cancel")}</button>
              {editMode
                ? <button className="btn sm" disabled={busy} onClick={() => setDraft([...draft, blankRow(nextNo(draft))])}>{tr("+ Row")}</button>
                : <button className="btn sm" disabled={busy} onClick={() => setEditMode(true)}>{tr("Edit")}</button>}
              {problem && <span className="pl-err">{tr(problem)}</span>}
            </>
          ) : (
            <>
              <button className="btn sm" disabled={busy} onClick={() => { setDraft(list.rows.length ? list.rows : [blankRow("01")]); setEditMode(true); }}>{tr("Edit")}</button>
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
        <JigSummary rows={rows} />
        {draft && match && <MatchNotice match={match} />}
        {rows.length === 0 && !draft ? (
          <div className="muted">{tr("No processes yet. Import the JIG / Machine / Tester list workbook of the model, or Edit to type them in.")}</div>
        ) : (
          <div className="pl-wrap">
            <table className="wk-it pl-t">
              <thead>
                <tr>
                  <th>No.</th><th className="l">{tr("Process")}</th><th className="l">{tr("Machine")}</th>
                  <th className="l">JIG</th><th className="l">{tr("Drawing No")}</th><th className="pl-narrow">{tr("Ver.")}</th><th className="pl-narrow">{tr("Qty")}</th>
                  {editMode && <th></th>}
                  <th className="l">{tr("Tester")}</th><th className="l">{tr("Material")}</th><th className="l">{tr("Note")}</th>
                  {editMode && <th></th>}
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => (
                  <ProcessRows key={draft ? i : r.id || i} row={r} edit={editMode}
                    onChange={(n) => draft && setDraft(draft.map((x, j) => (j === i ? n : x)))}
                    onRemove={() => draft && setDraft(draft.filter((_, j) => j !== i))} />
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </section>
  );
}
