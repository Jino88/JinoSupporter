import { Fragment, useEffect, useState, type ClipboardEvent } from "react";
import { workingApi } from "./api";
import {
  blankEntry, blankLog, entryProblem, fieldsOf, isLate, isLogLate, isOpenIssue, issueFields, issuesBySeq, logFieldsOf, logProblem,
  notesInOrder, pastedImages, suggestions,
} from "./logic";
import { ISSUE_CAUSES, ISSUE_PROCESSES, ISSUE_TYPES, PHASES, type Entry, type EntryFields, type EntryLog, type LogFields, type WorkingImage } from "./types";

type Filter = "open" | "closed" | "all";
const COLS = 14;

interface Lists { types: string[]; processes: string[]; causes: string[] }

/** Pictures being edited: new pasted files and the ids of saved pictures to remove on save. */
interface Pics { files: File[]; removed: string[] }

function usePictures(saved: WorkingImage[]) {
  const [files, setFiles] = useState<File[]>([]);
  const [removed, setRemoved] = useState<string[]>([]);
  const [problem, setProblem] = useState("");
  const onPaste = (ev: ClipboardEvent) => {
    const r = pastedImages(Array.from(ev.clipboardData.files), files.length);
    if (!r.files.length && !r.problem) return; // plain text paste
    ev.preventDefault();
    setProblem(r.problem);
    if (r.files.length) setFiles((cur) => [...cur, ...r.files]);
  };
  return {
    files, removed, problem, onPaste,
    kept: saved.filter((i) => !removed.includes(i.id)),
    drop: (index: number) => setFiles((cur) => cur.filter((_, j) => j !== index)),
    remove: (id: string) => setRemoved((cur) => [...cur, id]),
    reset: () => { setFiles([]); setRemoved([]); setProblem(""); },
  };
}

function PendingThumb(props: { file: File; onRemove: () => void }) {
  const [url, setUrl] = useState("");
  useEffect(() => {
    const u = URL.createObjectURL(props.file);
    setUrl(u);
    return () => URL.revokeObjectURL(u);
  }, [props.file]);
  return (
    <span className="wk-thumb new" title="Saved with the row">
      {url && <img src={url} alt={props.file.name} />}
      <button className="wk-thumb-x" aria-label="Remove picture" onClick={props.onRemove}>×</button>
    </span>
  );
}

/** Saved pictures (click to enlarge) and, while editing, the pasted ones not saved yet. */
function Thumbs(props: {
  images: WorkingImage[]; onOpen: (i: WorkingImage) => void;
  onRemove?: (id: string) => void; pending?: File[]; onDrop?: (index: number) => void;
}) {
  const pending = props.pending ?? [];
  if (!props.images.length && !pending.length) return null;
  return (
    <div className="wk-thumbs" onClick={(ev) => ev.stopPropagation()}>
      {props.images.map((i) => (
        <span key={i.id} className="wk-thumb">
          <button className="wk-thumb-img" title={i.fileName} onClick={() => props.onOpen(i)}>
            <img loading="lazy" src={workingApi.imageUrl(i.id)} alt={i.fileName} />
          </button>
          {props.onRemove && <button className="wk-thumb-x" aria-label="Remove picture" onClick={() => props.onRemove!(i.id)}>×</button>}
        </span>
      ))}
      {pending.map((f, j) => <PendingThumb key={j} file={f} onRemove={() => props.onDrop?.(j)} />)}
    </div>
  );
}

function IssueForm(props: {
  value: EntryFields; images: WorkingImage[]; lists: Lists; busy: boolean;
  onSave: (f: EntryFields, pics: Pics) => void; onCancel: () => void; onOpen: (i: WorkingImage) => void;
}) {
  const [f, setF] = useState(props.value);
  const pics = usePictures(props.images);
  const set = <K extends keyof EntryFields>(k: K, v: EntryFields[K]) => setF((cur) => ({ ...cur, [k]: v }));
  const problem = entryProblem(f);
  const text = (k: "category" | "process" | "cause" | "owner", label: string, list?: string) => (
    <div className="wk-f"><label>{label}</label>
      <input className="inp" value={f[k]} list={list} onChange={(e) => set(k, e.target.value)} /></div>
  );
  const date = (k: "date" | "due" | "doneDate", label: string) => (
    <div className="wk-f"><label>{label}</label>
      <input type="date" className="inp" value={f[k]} onChange={(e) => set(k, e.target.value)} /></div>
  );
  return (
    <div className="wk-form wk-ef wk-if" onPaste={pics.onPaste}>
      <div className="wk-f"><label>Dev phase</label>
        <select className="inp" value={f.phase} onChange={(e) => set("phase", e.target.value)}>
          <option value="">—</option>
          {PHASES.map((p) => <option key={p}>{p}</option>)}
          {f.phase && !PHASES.includes(f.phase as (typeof PHASES)[number]) && <option>{f.phase}</option>}
        </select>
      </div>
      {text("category", "Issue type", "wk-l-type")}
      {text("process", "Process", "wk-l-proc")}
      {text("cause", "Cause group", "wk-l-cause")}
      <div className="wk-f wide"><label>Issue</label>
        <input className="inp" value={f.title} maxLength={200} autoFocus onChange={(e) => set("title", e.target.value)} /></div>
      <div className="wk-f wide"><label>Pictures</label>
        <div className="wk-paste" tabIndex={0}>
          <span className="muted small">Click here and press Ctrl+V to paste a picture (PNG, JPEG, GIF, WebP, BMP · up to 10 MB).</span>
          <Thumbs images={pics.kept} onOpen={props.onOpen} onRemove={pics.remove} pending={pics.files} onDrop={pics.drop} />
        </div>
        {pics.problem && <span className="pi-dirty">{pics.problem}</span>}
      </div>
      {text("owner", "Owner")}
      {date("date", "Registered")}
      {date("due", "Due")}
      {date("doneDate", "Done")}
      <div className="wk-f wide"><label>Final result</label>
        <textarea className="inp wk-ta" rows={2} value={f.finalResult} onChange={(e) => set("finalResult", e.target.value)} /></div>
      <div className="wk-f wide wk-ef-act">
        {problem && f.title && <span className="pi-dirty">{problem}</span>}
        <button className="btn sm" onClick={props.onCancel} disabled={props.busy}>Cancel</button>
        <button className="btn sm primary" disabled={!!problem || props.busy}
          onClick={() => props.onSave(issueFields(f), { files: pics.files, removed: pics.removed })}>{props.busy ? "Saving…" : "Save"}</button>
      </div>
      <datalist id="wk-l-type">{props.lists.types.map((v) => <option key={v} value={v} />)}</datalist>
      <datalist id="wk-l-proc">{props.lists.processes.map((v) => <option key={v} value={v} />)}</datalist>
      <datalist id="wk-l-cause">{props.lists.causes.map((v) => <option key={v} value={v} />)}</datalist>
    </div>
  );
}

/** Done / not done of a note or issue: the done date, or Open. */
const DoneCell = (props: { doneDate: string }) =>
  props.doneDate ? <span className="wk-st closed">{props.doneDate}</span> : <span className="wk-st open">Open</span>;

/** A progress note shown in the issue's columns: text + pictures under Issue, then owner and the dates. */
function LogRow(props: { no: string; log: EntryLog; today: string; busy: boolean; onEdit: () => void; onDelete: () => void; onOpen: (i: WorkingImage) => void }) {
  const { log } = props;
  const late = isLogLate(logFieldsOf(log), props.today);
  return (
    <tr className={"wk-it-log" + (log.doneDate ? " done" : "")}>
      <td className="tog">↳</td>
      <td className="n sub">{props.no}</td>
      <td colSpan={4}></td>
      <td className="l wk-it-title"><div className="wk-text">{log.text}</div><Thumbs images={log.images ?? []} onOpen={props.onOpen} /></td>
      <td>{log.owner || <span className="muted" title="Written by">{log.createdBy}</span>}</td>
      <td>{log.date}</td>
      <td className={late ? "late" : ""}>{log.due}{late && <span className="wk-late">Late</span>}</td>
      <td><DoneCell doneDate={log.doneDate ?? ""} /></td>
      <td></td>
      <td></td>
      <td className="act">
        <button className="wk-link" onClick={props.onEdit} disabled={props.busy}>Edit</button>
        <button className="wk-link danger" onClick={props.onDelete} disabled={props.busy}>Delete</button>
      </td>
    </tr>
  );
}

/** Add / edit a progress note in place, in the same columns as the issue. Ctrl+V in the text pastes pictures. */
function LogEditRow(props: {
  no: string; value: LogFields; images: WorkingImage[]; today: string; busy: boolean; isNew: boolean;
  onSave: (f: LogFields, pics: Pics) => Promise<boolean>; onCancel?: () => void; onOpen: (i: WorkingImage) => void;
}) {
  const [f, setF] = useState(props.value);
  const pics = usePictures(props.images);
  const set = <K extends keyof LogFields>(k: K, v: LogFields[K]) => setF((cur) => ({ ...cur, [k]: v }));
  const problem = logProblem(f);
  const save = async () => {
    if (await props.onSave(f, { files: pics.files, removed: pics.removed }) && props.isNew) {
      setF(props.value);
      pics.reset();
    }
  };
  return (
    <tr className={"wk-it-log edit" + (props.isNew ? " new" : "")}>
      <td className="tog">{props.isNew ? "+" : "↳"}</td>
      <td className="n sub">{props.no}</td>
      <td colSpan={4}></td>
      <td className="l wk-it-title">
        <textarea className="inp wk-ta" rows={2} value={f.text} onPaste={pics.onPaste} autoFocus={!props.isNew}
          placeholder={props.isNew ? "Add today's progress… (Ctrl+V pastes a picture)" : "Ctrl+V pastes a picture"}
          onChange={(e) => set("text", e.target.value)} />
        <Thumbs images={pics.kept} onOpen={props.onOpen} onRemove={pics.remove} pending={pics.files} onDrop={pics.drop} />
        {(pics.problem || (f.text && problem)) && <div className="pi-dirty">{pics.problem || problem}</div>}
      </td>
      <td><input className="inp sm wk-in-owner" value={f.owner} placeholder="Owner" onChange={(e) => set("owner", e.target.value)} /></td>
      <td><input type="date" className="inp sm" aria-label="Note date" value={f.date} onChange={(e) => set("date", e.target.value)} /></td>
      <td><input type="date" className="inp sm" aria-label="Due" value={f.due} onChange={(e) => set("due", e.target.value)} /></td>
      <td>
        <label className="wk-done">
          <input type="checkbox" checked={!!f.doneDate}
            onChange={(e) => set("doneDate", e.target.checked ? (f.date > props.today ? f.date : props.today) : "")} /> Done
        </label>
        {f.doneDate && <input type="date" className="inp sm" aria-label="Done date" value={f.doneDate} onChange={(e) => set("doneDate", e.target.value)} />}
      </td>
      <td></td>
      <td></td>
      <td className="act">
        {!props.isNew && <button className="btn sm" onClick={props.onCancel} disabled={props.busy}>Cancel</button>}
        <button className="btn sm primary" onClick={save} disabled={props.busy || !!problem}>{props.isNew ? "Add" : "Save"}</button>
      </td>
    </tr>
  );
}

/**
 * Issue register: No., dev phase, classification (issue type / process / cause group), issue, owner,
 * registered / due / done dates and final result. An issue opens into its progress notes, one row each in the same columns.
 */
export default function IssueTable(props: {
  modelName: string;
  modelPhase: string;
  entries: Entry[];
  today: string;
  onChanged: (list: Entry[]) => void;
  onError: (e: unknown) => void;
}) {
  const { modelName, entries, today } = props;
  const [filter, setFilter] = useState<Filter>("open");
  const [editing, setEditing] = useState<string | null>(null); // issue id, "new" or null
  const [editingLog, setEditingLog] = useState<string | null>(null);
  const [open, setOpen] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [viewer, setViewer] = useState<WorkingImage | null>(null);

  const all = issuesBySeq(entries);
  const shown = all.filter((e) => filter === "all" || (filter === "open") === isOpenIssue(e));
  const lists: Lists = {
    types: suggestions(all.map((e) => e.category), ISSUE_TYPES),
    processes: suggestions(all.map((e) => e.process), ISSUE_PROCESSES),
    causes: suggestions(all.map((e) => e.cause), ISSUE_CAUSES),
  };

  /**
   * Removes / uploads the pictures of a saved row and returns its pictures afterwards. A failure is reported
   * without undoing the saved row, so pressing Save again does not create it twice.
   */
  const syncPictures = async (entryId: string, logId: string, saved: WorkingImage[], pics: Pics) => {
    let kept = saved;
    try {
      for (const id of pics.removed) {
        await workingApi.deleteImage(modelName, entryId, id);
        kept = kept.filter((i) => i.id !== id);
      }
      return pics.files.length ? [...kept, ...await workingApi.addImages(modelName, entryId, logId, pics.files)] : kept;
    } catch (e) {
      props.onError(e);
      return kept;
    }
  };

  const run = async (work: () => Promise<Entry[]>) => {
    setBusy(true);
    try { props.onChanged(await work()); setEditing(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const add = (f: EntryFields, pics: Pics) => run(async () => {
    const saved = await workingApi.addEntry(modelName, f);
    return [...entries, { ...saved, images: await syncPictures(saved.id, "", saved.images ?? [], pics) }];
  });
  const save = (e: Entry, f: EntryFields, pics: Pics) => run(async () => {
    const saved = await workingApi.saveEntry(modelName, e.id, f, e.revision);
    const images = await syncPictures(e.id, "", saved.images ?? e.images ?? [], pics);
    return entries.map((x) => (x.id === e.id ? { ...saved, images } : x));
  });
  const remove = (e: Entry) => {
    if (!window.confirm(`Delete issue No. ${e.seq} "${e.title}" and its ${e.logs?.length ?? 0} notes?`)) return;
    void run(async () => { await workingApi.deleteEntry(modelName, e.id); return entries.filter((x) => x.id !== e.id); });
  };

  const setLogs = (e: Entry, change: (logs: EntryLog[]) => EntryLog[]) =>
    props.onChanged(entries.map((x) => (x.id === e.id ? { ...x, logs: change(x.logs ?? []) } : x)));
  const runLog = async (work: () => Promise<void>) => {
    setBusy(true);
    try { await work(); return true; } catch (err) { props.onError(err); return false; } finally { setBusy(false); }
  };
  const addLog = (e: Entry, f: LogFields, pics: Pics) => runLog(async () => {
    const saved = await workingApi.addLog(modelName, e.id, f);
    const images = await syncPictures(e.id, saved.id, saved.images ?? [], pics);
    setLogs(e, (logs) => [...logs, { ...saved, images }]);
  });
  const saveLog = (e: Entry, l: EntryLog, f: LogFields, pics: Pics) => runLog(async () => {
    const saved = await workingApi.saveLog(modelName, e.id, l, f);
    const images = await syncPictures(e.id, l.id, saved.images ?? l.images ?? [], pics);
    setLogs(e, (logs) => logs.map((x) => (x.id === l.id ? { ...saved, images } : x)));
    setEditingLog(null);
  });
  const removeLog = (e: Entry, l: EntryLog) => {
    if (!window.confirm(`Delete the note of ${l.date}?`)) return;
    void runLog(async () => {
      await workingApi.deleteLog(modelName, e.id, l.id);
      setLogs(e, (logs) => logs.filter((x) => x.id !== l.id));
    });
  };

  const toggle = (id: string) => setOpen((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });
  const count = (f: Filter) => all.filter((e) => f === "all" || (f === "open") === isOpenIssue(e)).length;

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Issues <span className="cnt">{shown.length}</span></h2>
        <div className="pi-legend">
          {(["open", "closed", "all"] as Filter[]).map((v) => (
            <button key={v} className={"wk-chip" + (filter === v ? " on" : "")} onClick={() => setFilter(v)}>
              {v[0].toUpperCase() + v.slice(1)} {count(v)}
            </button>
          ))}
          <button className="btn sm primary" disabled={editing === "new"} onClick={() => setEditing("new")}>+ New issue</button>
        </div>
      </div>
      {editing === "new" && (
        <div className="pi-sec-b">
          <IssueForm value={blankEntry("issue", today, props.modelPhase)} images={[]} lists={lists} busy={busy}
            onSave={add} onCancel={() => setEditing(null)} onOpen={setViewer} />
        </div>
      )}
      <div className="pi-wrap">
        <table className="wk-it">
          <thead>
            <tr>
              <th rowSpan={2}></th><th rowSpan={2}>No.</th><th rowSpan={2}>Dev phase</th>
              <th colSpan={3} className="grp">Classification</th>
              <th rowSpan={2} className="l">Issue</th><th rowSpan={2}>Owner</th>
              <th rowSpan={2}>Registered</th><th rowSpan={2}>Due</th><th rowSpan={2}>Done</th>
              <th rowSpan={2} className="l">Final result</th><th rowSpan={2}>Notes</th><th rowSpan={2}></th>
            </tr>
            <tr><th>Issue type</th><th>Process</th><th>Cause group</th></tr>
          </thead>
          <tbody>
            {!shown.length && (
              <tr><td colSpan={COLS} className="wk-it-empty">{all.length ? "No issue in this filter." : "No issue yet. Add one with + New issue."}</td></tr>
            )}
            {shown.map((e) => {
              const expanded = open.has(e.id);
              const notes = notesInOrder(e.logs ?? []);
              const last = notes.at(-1);
              return (
                <Fragment key={e.id}>
                  <tr className={"wk-it-row" + (isOpenIssue(e) ? "" : " done") + (expanded ? " on" : "")} onClick={() => toggle(e.id)}>
                    <td className="tog">{expanded ? "▾" : "▸"}</td>
                    <td className="n">{e.seq || "—"}</td>
                    <td>{e.phase || "—"}</td>
                    <td>{e.category}</td>
                    <td>{e.process}</td>
                    <td>{e.cause}</td>
                    <td className="l wk-it-title" title={e.body || undefined}>
                      <b>{e.title}</b>
                      <Thumbs images={e.images ?? []} onOpen={setViewer} />
                      {last && !expanded && <div className="muted small">{last.date.slice(5)} · {last.text.split("\n")[0]}</div>}
                    </td>
                    <td>{e.owner}</td>
                    <td>{e.date}</td>
                    <td className={isLate(e, today) ? "late" : ""}>{e.due}{isLate(e, today) && <span className="wk-late">Late</span>}</td>
                    <td><DoneCell doneDate={e.doneDate} /></td>
                    <td className="l wk-it-final">{e.finalResult}</td>
                    <td className="n">{notes.length}</td>
                    <td className="act" onClick={(ev) => ev.stopPropagation()}>
                      <button className="wk-link" onClick={() => setEditing(e.id)} disabled={busy}>Edit</button>
                      <button className="wk-link danger" onClick={() => remove(e)} disabled={busy}>Delete</button>
                    </td>
                  </tr>
                  {editing === e.id && (
                    <tr className="wk-it-sub"><td colSpan={COLS}>
                      <IssueForm value={fieldsOf(e)} images={e.images ?? []} lists={lists} busy={busy}
                        onSave={(f, pics) => save(e, f, pics)} onCancel={() => setEditing(null)} onOpen={setViewer} />
                    </td></tr>
                  )}
                  {expanded && editing !== e.id && notes.map((l, i) => editingLog === l.id ? (
                    <LogEditRow key={l.id} no={`${e.seq}-${i + 1}`} value={logFieldsOf(l)} images={l.images ?? []} today={today} busy={busy} isNew={false}
                      onSave={(f, pics) => saveLog(e, l, f, pics)} onCancel={() => setEditingLog(null)} onOpen={setViewer} />
                  ) : (
                    <LogRow key={l.id} no={`${e.seq}-${i + 1}`} log={l} today={today} busy={busy}
                      onEdit={() => setEditingLog(l.id)} onDelete={() => removeLog(e, l)} onOpen={setViewer} />
                  ))}
                  {expanded && editing !== e.id && (
                    <LogEditRow key="new" no="" value={blankLog(today)} images={[]} today={today} busy={busy} isNew
                      onSave={(f, pics) => addLog(e, f, pics)} onOpen={setViewer} />
                  )}
                </Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
      {viewer && (
        <div className="wk-viewer" role="dialog" aria-label="Picture" onClick={() => setViewer(null)}>
          <div onClick={(ev) => ev.stopPropagation()}>
            <div className="wk-viewer-h">
              <span>{viewer.fileName} · {viewer.createdBy}</span>
              <a href={workingApi.imageUrl(viewer.id)} target="_blank" rel="noreferrer">Open original</a>
              <button className="btn sm" onClick={() => setViewer(null)}>Close</button>
            </div>
            <img src={workingApi.imageUrl(viewer.id)} alt={viewer.fileName} />
          </div>
        </div>
      )}
    </section>
  );
}
