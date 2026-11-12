import { useState } from "react";
import DateBox from "./DateBox";
import { useTr } from "./i18n";
import { workingApi } from "./api";
import { blankEntry, byMonth, entryProblem, fieldsOf, isLate, isOpenIssue } from "./logic";
import { KIND_LABEL, RESULTS, type Entry, type EntryFields, type EntryKind } from "./types";

const TITLES: Record<EntryKind, string> = { dev: "Development history", issue: "Issues", review: "Review history" };
const ADD: Record<EntryKind, string> = { dev: "+ Add entry", issue: "+ New issue", review: "+ Add review" };
const OWNER: Record<EntryKind, string> = { dev: "By", issue: "Owner", review: "Chair" };
const KINDS: EntryKind[] = ["dev", "issue", "review"];
/** Issues are added and edited in the Issues tab (IssueTable). */
const ADDABLE: EntryKind[] = ["dev", "review"];

type IssueFilter = "all" | "open" | "closed";

function EntryForm(props: { value: EntryFields; kindFixed: boolean; busy: boolean; onSave: (f: EntryFields) => void; onCancel: () => void }) {
  const [f, setF] = useState(props.value);
  const set = <K extends keyof EntryFields>(k: K, v: EntryFields[K]) => setF((cur) => ({ ...cur, [k]: v }));
  const problem = entryProblem(f);
  return (
    <div className="wk-form wk-ef">
      {!props.kindFixed && (
        <div className="wk-f"><label>Type</label>
          <select className="inp" value={f.kind} onChange={(e) => {
            const kind = e.target.value as EntryKind;
            setF({ ...blankEntry(kind, f.date), title: f.title, body: f.body, owner: f.owner });
          }}>
            {ADDABLE.map((k) => <option key={k} value={k}>{KIND_LABEL[k]}</option>)}
          </select>
        </div>
      )}
      <div className="wk-f"><label>Date</label><DateBox label="Date" value={f.date} onChange={(v) => set("date", v)} /></div>
      {f.kind === "review" && (
        <div className="wk-f"><label>Result</label>
          <select className="inp" value={f.result} onChange={(e) => set("result", e.target.value)}>{RESULTS.map((s) => <option key={s}>{s}</option>)}</select></div>
      )}
      <div className="wk-f"><label>{OWNER[f.kind]}</label><input className="inp" value={f.owner} onChange={(e) => set("owner", e.target.value)} /></div>
      <div className="wk-f wide"><label>Title</label><input className="inp" value={f.title} maxLength={200} autoFocus onChange={(e) => set("title", e.target.value)} /></div>
      <div className="wk-f wide"><label>Details</label><textarea className="inp wk-ta" rows={5} value={f.body} onChange={(e) => set("body", e.target.value)} /></div>
      <div className="wk-f wide wk-ef-act">
        {problem && f.title && <span className="pi-dirty">{problem}</span>}
        <button className="btn sm" onClick={props.onCancel} disabled={props.busy}>Cancel</button>
        <button className="btn sm primary" disabled={!!problem || props.busy} onClick={() => props.onSave(f)}>{props.busy ? "Saving…" : "Save"}</button>
      </div>
    </div>
  );
}

function Badges({ e, today }: { e: Entry; today: string }) {
  return (
    <>
      <span className={"wk-kind " + e.kind}>{KIND_LABEL[e.kind]}</span>
      {e.kind === "issue" && e.seq > 0 && <span className="wk-tag">No. {e.seq}</span>}
      {e.kind === "issue" && e.severity && <span className={"wk-sev " + e.severity.toLowerCase()}>{e.severity}</span>}
      {e.kind === "issue" && <span className={"wk-st " + (isOpenIssue(e) ? "open" : "closed")}>{isOpenIssue(e) ? "Open" : "Closed"}</span>}
      {e.kind === "review" && e.result && <span className={"wk-res " + e.result.toLowerCase()}>{e.result}</span>}
      {isLate(e, today) && <span className="wk-late">Past due</span>}
    </>
  );
}

/** Entries of one kind (or all kinds when `kind` is null) as a dated timeline with add / edit / delete. */
export default function EntryList(props: {
  modelName: string;
  kind: EntryKind | null;
  entries: Entry[];
  today: string;
  onChanged: (list: Entry[]) => void;
  onError: (e: unknown) => void;
  onOpenIssues: (issue: Entry) => void;
}) {
  const tr = useTr();
  const { modelName, kind, entries, today } = props;
  const [editing, setEditing] = useState<string | null>(null); // entry id, "new" or null
  const [busy, setBusy] = useState(false);
  const [issueFilter, setIssueFilter] = useState<IssueFilter>("open");
  const [kinds, setKinds] = useState<Set<EntryKind>>(new Set(KINDS));

  let list = entries.filter((e) => (kind ? e.kind === kind : kinds.has(e.kind)));
  if (kind === "issue" && issueFilter !== "all") list = list.filter((e) => (issueFilter === "open") === isOpenIssue(e));

  const run = async (work: () => Promise<Entry[]>) => {
    setBusy(true);
    try { props.onChanged(await work()); setEditing(null); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const add = (f: EntryFields) => run(async () => [...entries, await workingApi.addEntry(modelName, f)]);
  const save = (e: Entry, f: EntryFields) =>
    run(async () => { const saved = await workingApi.saveEntry(modelName, e.id, f, e.revision); return entries.map((x) => (x.id === e.id ? saved : x)); });
  const remove = (e: Entry) => {
    if (!window.confirm(`Delete "${e.title}"?`)) return;
    run(async () => { await workingApi.deleteEntry(modelName, e.id); return entries.filter((x) => x.id !== e.id); });
  };
  const toggleKind = (k: EntryKind) => setKinds((s) => { const n = new Set(s); if (n.has(k)) n.delete(k); else n.add(k); return n; });

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>{tr(kind ? TITLES[kind] : "Timeline")} <span className="cnt">{list.length}</span></h2>
        <div className="pi-legend">
          {kind === "issue" && (["open", "closed", "all"] as IssueFilter[]).map((v) => (
            <button key={v} className={"wk-chip" + (issueFilter === v ? " on" : "")} onClick={() => setIssueFilter(v)}>{v[0].toUpperCase() + v.slice(1)}</button>
          ))}
          {!kind && KINDS.map((k) => (
            <button key={k} className={"wk-chip " + k + (kinds.has(k) ? " on" : "")} onClick={() => toggleKind(k)}>{KIND_LABEL[k]}</button>
          ))}
          <button className="btn sm primary" disabled={editing === "new"} onClick={() => setEditing("new")}>{kind ? ADD[kind] : "+ Add"}</button>
        </div>
      </div>
      <div className="pi-sec-b">
        {editing === "new" && (
          <EntryForm value={blankEntry(kind ?? "dev", today)} kindFixed={!!kind} busy={busy} onSave={add} onCancel={() => setEditing(null)} />
        )}
        {!list.length && editing !== "new" && <div className="pi-empty">Nothing here yet.</div>}
        <ol className="wk-tl">
          {byMonth(list).map((m) => [
            <li key={m.month} className="wk-month">{m.month}</li>,
            ...m.entries.map((e) => (
              <li key={e.id} className={"wk-item " + e.kind}>
                <span className="wk-date">{e.date.slice(5)}</span>
                <span className="wk-dot" />
                {editing === e.id ? (
                  <EntryForm value={fieldsOf(e)} kindFixed busy={busy} onSave={(f) => save(e, f)} onCancel={() => setEditing(null)} />
                ) : (
                  <div className="wk-body">
                    <div className="wk-title"><Badges e={e} today={today} /><b>{tr(e.title)}</b></div>
                    {e.body && <div className="wk-text">{tr(e.body)}</div>}
                    <div className="wk-meta">
                      {e.owner && <span>{OWNER[e.kind]} {e.owner}</span>}
                      {e.kind === "issue" && e.due && <span>Due {e.due}</span>}
                      <span>Written by {e.createdBy || "—"}{e.updatedAt !== e.createdAt && e.updatedBy ? ` · edited by ${e.updatedBy}` : ""}</span>
                      {e.kind === "issue" ? (
                        <button className="wk-link" onClick={() => props.onOpenIssues(e)}>Open in Issues ({e.logs?.length ?? 0} notes)</button>
                      ) : (
                        <>
                          <button className="wk-link" onClick={() => setEditing(e.id)} disabled={busy}>Edit</button>
                          <button className="wk-link danger" onClick={() => remove(e)} disabled={busy}>Delete</button>
                        </>
                      )}
                    </div>
                  </div>
                )}
              </li>
            )),
          ])}
        </ol>
      </div>
    </section>
  );
}
