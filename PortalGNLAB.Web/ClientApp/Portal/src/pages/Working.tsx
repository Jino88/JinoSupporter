import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { localToday } from "./productionInput/logic";
import { workingApi } from "./working/api";
import EntryList from "./working/EntryList";
import InfoTab from "./working/InfoTab";
import IssueTable from "./working/IssueTable";
import { buildTree, countByKind, filterTree, firstModel, isOpenIssue, pathOf, sortEntries, type GroupNode } from "./working/logic";
import PlanView from "./working/PlanView";
import SetupPlan from "./working/SetupPlan";
import type { Entry, GroupsResponse, WorkingModel } from "./working/types";
import "./working/working.css";

type Tab = "info" | "plan" | "setup" | "dev" | "issue" | "review" | "all";
const TABS: [Tab, string][] = [
  ["info", "Model Info"], ["plan", "Dev Plan"], ["setup", "Setup Plan"], ["dev", "Dev History"], ["issue", "Issues"], ["review", "Reviews"], ["all", "Timeline"],
];
const isTab = (v: string | null): v is Tab => TABS.some(([k]) => k === v);

/** Keeps ?model=&tab= in the address so a model page can be linked and survives reload. */
function writeQuery(model: string | null, tab: Tab) {
  const q = new URLSearchParams(window.location.search);
  if (model) q.set("model", model); else q.delete("model");
  q.set("tab", tab);
  window.history.replaceState(null, "", `${window.location.pathname}?${q}`);
}

export default function Working() {
  const query = new URLSearchParams(window.location.search);
  const [groups, setGroups] = useState<GroupsResponse | null>(null);
  const [filter, setFilter] = useState("");
  const [closed, setClosed] = useState<Set<string>>(new Set());
  const [selected, setSelected] = useState<string | null>(query.get("model"));
  const [tab, setTab] = useState<Tab>(isTab(query.get("tab")) ? (query.get("tab") as Tab) : "info");
  const [model, setModel] = useState<WorkingModel | null>(null);
  const [entries, setEntries] = useState<Entry[]>([]);
  const [infoDirty, setInfoDirty] = useState(false);
  const [error, setError] = useState("");
  const load = useRef(0);
  const today = localToday();

  const tree = useMemo(() => buildTree(groups?.groups ?? []), [groups]);
  const shown = useMemo(() => filterTree(tree, filter), [tree, filter]);
  const fail = (e: unknown) => setError((e as Error).message);

  useEffect(() => {
    workingApi.groups()
      .then((g) => {
        setGroups(g);
        setSelected((cur) => cur ?? firstModel(buildTree(g.groups)));
      })
      .catch(fail);
  }, []);

  const reload = useCallback((name: string) => {
    const ticket = ++load.current;
    Promise.all([workingApi.model(name), workingApi.entries(name)])
      .then(([m, list]) => {
        if (ticket !== load.current) return;
        setModel(m);
        setEntries(sortEntries(list));
        setError("");
      })
      .catch(fail);
  }, []);

  useEffect(() => {
    writeQuery(selected, tab);
  }, [selected, tab]);

  useEffect(() => {
    setModel(null);
    setEntries([]);
    setInfoDirty(false);
    if (selected) reload(selected);
  }, [selected, reload]);

  const pick = (name: string) => {
    if (name === selected) return;
    if (infoDirty && !window.confirm("Discard the unsaved model info?")) return;
    setSelected(name);
  };
  const toggle = (id: string) => setClosed((s) => { const n = new Set(s); if (n.has(id)) n.delete(id); else n.add(id); return n; });

  const counts = countByKind(entries);
  const openIssues = entries.filter(isOpenIssue).length;
  const path = selected ? pathOf(tree, selected) : [];
  const onChanged = (list: Entry[]) => setEntries(sortEntries(list));

  const renderNode = (n: GroupNode, depth: number) => (
    <div key={n.id} className={`wk-g d${Math.min(depth, 2)}`}>
      <button className="wk-gh" onClick={() => toggle(n.id)} aria-expanded={!closed.has(n.id)}>
        <span className="wk-caret">{closed.has(n.id) && !filter ? "▸" : "▾"}</span>{n.name}<span className="wk-cnt">{n.total}</span>
      </button>
      {(!closed.has(n.id) || filter) && (
        <>
          {n.children.map((c) => renderNode(c, depth + 1))}
          {n.models.map((m) => (
            <button key={m} className={"wk-m" + (m === selected ? " on" : "")} onClick={() => pick(m)}>{m}</button>
          ))}
        </>
      )}
    </div>
  );

  return (
    <div className="tools-page pi-page wk-page">
      <header className="page-head">
        <h1>Working</h1>
        <p className="muted">Everything about one model in one place: model info, the development plan from Production Input, and the development, issue and review history.</p>
      </header>
      {error && <div className="alert" role="alert">{error}</div>}

      <div className="wk-lay">
        <aside className="card wk-side">
          <div className="wk-side-h">
            <h2>Model groups</h2>
            <input className="inp sm" placeholder="Search model…" value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Search model" />
          </div>
          <div className="wk-tree">
            {groups === null && <div className="muted small">Loading…</div>}
            {groups?.error && <div className="wk-warn">{groups.error}</div>}
            {groups && !groups.error && !tree.length && <div className="muted small">No model group yet. Add groups in the Company Glossary.</div>}
            {shown.map((n) => renderNode(n, 0))}
            {filter && tree.length > 0 && !shown.length && <div className="muted small">No model matches “{filter}”.</div>}
          </div>
          {groups?.source && <div className="wk-src" title={groups.source}>Source: Company Glossary</div>}
        </aside>

        <section className="wk-main">
          {!selected ? (
            <div className="card pi-empty">Pick a model on the left.</div>
          ) : (
            <>
              <div className="wk-head">
                <div>
                  <div className="wk-crumb">{path.length ? path.join(" › ") : "Not in a group"}</div>
                  <h2>{selected}</h2>
                  <div className="muted small">
                    {[model?.info.customer, model?.info.build, model?.info.owner && `PM ${model.info.owner}`].filter(Boolean).join(" · ") || "No model info yet"}
                  </div>
                </div>
                <div className="wk-head-r">
                  {model?.info.phase && <span className="wk-phase">{model.info.phase}</span>}
                  {openIssues > 0 && <span className="wk-open">{openIssues} open issue{openIssues > 1 ? "s" : ""}</span>}
                </div>
              </div>

              <div className="pi-tabs" role="tablist">
                {TABS.map(([k, label]) => {
                  const n = k === "dev" || k === "review" ? counts[k] : k === "issue" ? openIssues : k === "all" ? entries.length : null;
                  return (
                    <button key={k} role="tab" aria-selected={tab === k} className={tab === k ? "on" : ""}
                      onClick={() => {
                        if (k === tab) return;
                        if (infoDirty && !window.confirm("Discard the unsaved model info?")) return;
                        setInfoDirty(false);
                        setTab(k);
                      }}>{label}{n != null && n > 0 && <span className="cnt">{n}</span>}</button>
                  );
                })}
              </div>

              {!model ? <div className="muted">Loading…</div> : (
                <>
                  {tab === "info" && (
                    <InfoTab key={selected} model={model} entries={entries} today={today} onDirty={setInfoDirty}
                      onSaved={(m) => { setModel(m); setInfoDirty(false); }} onError={fail} />
                  )}
                  {tab === "plan" && <PlanView key={selected} modelName={selected} />}
                  {tab === "setup" && (
                    <SetupPlan key={selected} modelName={selected} modelPhase={model.info.phase ?? ""}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} />
                  )}
                  {tab === "issue" && (
                    <IssueTable key={selected} modelName={selected} modelPhase={model.info.phase ?? ""}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} />
                  )}
                  {(tab === "dev" || tab === "review" || tab === "all") && (
                    <EntryList key={selected + tab} modelName={selected} kind={tab === "all" ? null : tab}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} onOpenIssues={() => setTab("issue")} />
                  )}
                </>
              )}
            </>
          )}
        </section>
      </div>
    </div>
  );
}
