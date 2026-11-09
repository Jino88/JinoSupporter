import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { localToday } from "./productionInput/logic";
import { workingApi } from "./working/api";
import EntryList from "./working/EntryList";
import InfoTab from "./working/InfoTab";
import IssueTable from "./working/IssueTable";
import {
  buildTree, countByKind, filterTree, firstModel, isOpenIssue, issuesOfStage, pathOf, replaceSome, sortEntries, type GroupNode, type IssueStage,
} from "./working/logic";
import PlanView from "./working/PlanView";
import SetupPlan from "./working/SetupPlan";
import { bmesApi, type BmesMaterial } from "./working/bmes";
import { hideModels, treeApi, withAdded, withBmes, type TreeSettings } from "./working/tree";
import TreeManager from "./working/TreeManager";
import { phaseName, type Entry, type GroupsResponse, type WorkingModel } from "./working/types";
import "./working/working.css";

type Tab = "info" | "plan" | "devissue" | "dev" | "review" | "all" | "setup" | "issue";
/** Categories of a model page and the tabs under each (Model Info has none). */
const CATEGORIES: { label: string; tabs: [Tab, string][] }[] = [
  { label: "Model Info", tabs: [["info", "Model Info"]] },
  { label: "Development", tabs: [["plan", "Dev Plan"], ["devissue", "Dev Issue"], ["dev", "Dev History"], ["review", "Reviews"], ["all", "Timeline"]] },
  { label: "Mass Production", tabs: [["setup", "Setup Plan"], ["issue", "Issues"]] },
];
const categoryOf = (tab: Tab) => CATEGORIES.find((c) => c.tabs.some(([k]) => k === tab))!;
const isTab = (v: string | null): v is Tab => CATEGORIES.some((c) => c.tabs.some(([k]) => k === v));
/** The issue list of a tab: Dev Issue (development) or Issues (mass production). */
const stageOf = (tab: Tab): IssueStage | null => (tab === "devissue" ? "dev" : tab === "issue" ? "mp" : null);

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
  const [treeSettings, setTreeSettings] = useState<TreeSettings>({ hidden: [], added: [] });
  const [bmes, setBmes] = useState<BmesMaterial[]>([]);
  const [managing, setManaging] = useState(false);
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

  // every model (Glossary + BMES Make Model Group + added here), then the ones not hidden in Manage
  const merged = useMemo(() => withBmes(groups?.groups ?? [], bmes, treeSettings.added), [groups, bmes, treeSettings.added]);
  const fullTree = useMemo(
    () => buildTree(withAdded(merged.groups, treeSettings.added, treeSettings.groups, treeSettings.placed)), [merged, treeSettings]);
  const tree = useMemo(() => hideModels(fullTree, treeSettings.hidden), [fullTree, treeSettings.hidden]);
  const shown = useMemo(() => filterTree(tree, filter), [tree, filter]);
  const fail = (e: unknown) => setError((e as Error).message);

  useEffect(() => {
    // without the tree settings (older server, or a failure) every Glossary model is shown
    const empty: TreeSettings = { hidden: [], added: [] };
    Promise.all([workingApi.groups(), treeApi.get().catch(() => empty), bmesApi.materials().catch(() => [] as BmesMaterial[])])
      .then(([g, t, b]) => {
        setGroups(g);
        setTreeSettings(t);
        setBmes(b);
        const base = withBmes(g.groups, b, t.added).groups;
        setSelected((cur) => cur ?? firstModel(hideModels(buildTree(withAdded(base, t.added, t.groups, t.placed)), t.hidden)));
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
  const openIn = (stage: IssueStage) => issuesOfStage(entries, stage).filter(isOpenIssue).length;
  const go = (k: Tab) => {
    if (k === tab) return;
    if (infoDirty && !window.confirm("Discard the unsaved model info?")) return;
    setInfoDirty(false);
    setTab(k);
  };
  const category = categoryOf(tab);
  const modelPhase = phaseName(model?.info.phase);
  const stage = stageOf(tab);
  const stageIssues = stage ? issuesOfStage(entries, stage) : [];

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
            <div className="wk-side-t">
              <h2>Model groups</h2>
              <button className="btn sm" disabled={!groups} onClick={() => setManaging(true)} title="Choose the models shown here and add models">Manage</button>
            </div>
            <input className="inp sm" placeholder="Search model…" value={filter} onChange={(e) => setFilter(e.target.value)} aria-label="Search model" />
          </div>
          <div className="wk-tree">
            {groups === null && <div className="muted small">Loading…</div>}
            {groups?.error && <div className="wk-warn">{groups.error}</div>}
            {groups && !groups.error && !fullTree.length && <div className="muted small">No model group yet. Add groups in the Company Glossary, or add a model with Manage.</div>}
            {fullTree.length > 0 && !tree.length && <div className="muted small">Every model is hidden. Choose the models to show with Manage.</div>}
            {shown.map((n) => renderNode(n, 0))}
            {filter && tree.length > 0 && !shown.length && <div className="muted small">No model matches “{filter}”.</div>}
          </div>
          {groups?.source && (
            <div className="wk-src" title={groups.source}>
              Source: Company Glossary{merged.bmesModels.length ? ` + ${merged.bmesModels.length} from BMES Model Group` : ""}
              {treeSettings.added.length ? ` + ${treeSettings.added.length} added` : ""}
              {treeSettings.hidden.length ? ` · ${treeSettings.hidden.length} hidden` : ""}
            </div>
          )}
        </aside>
        {managing && (
          <TreeManager tree={fullTree} bmesModels={merged.bmesModels} settings={treeSettings} onChanged={setTreeSettings} onClose={() => setManaging(false)} onError={fail} />
        )}

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
                  {modelPhase && <span className="wk-phase">{modelPhase}</span>}
                  {openIssues > 0 && <span className="wk-open">{openIssues} open issue{openIssues > 1 ? "s" : ""}</span>}
                </div>
              </div>

              <div className="pi-tabs" role="tablist" aria-label="Category">
                {CATEGORIES.map((c) => {
                  const n = c.label === "Development" ? openIn("dev") : c.label === "Mass Production" ? openIn("mp") : 0;
                  return (
                    <button key={c.label} role="tab" aria-selected={c === category} className={c === category ? "on" : ""}
                      title={n ? `${n} open issue${n > 1 ? "s" : ""}` : undefined} onClick={() => go(c.tabs[0][0])}>
                      {c.label}{n > 0 && <span className="cnt bad">{n}</span>}
                    </button>
                  );
                })}
              </div>
              {category.tabs.length > 1 && (
                <div className="wk-subtabs" role="tablist" aria-label={category.label}>
                  {category.tabs.map(([k, label]) => {
                    const n = k === "dev" || k === "review" ? counts[k] : k === "devissue" ? openIn("dev") : k === "issue" ? openIn("mp")
                      : k === "all" ? entries.length : null;
                    return (
                      <button key={k} role="tab" aria-selected={tab === k} className={"wk-chip" + (tab === k ? " on" : "")} onClick={() => go(k)}>
                        {label}{n != null && n > 0 && <span className="cnt">{n}</span>}
                      </button>
                    );
                  })}
                </div>
              )}

              {!model ? <div className="muted">Loading…</div> : (
                <>
                  {tab === "info" && (
                    <InfoTab key={selected} model={model} onDirty={setInfoDirty}
                      onSaved={(m) => { setModel(m); setInfoDirty(false); }} onError={fail} />
                  )}
                  {tab === "plan" && <PlanView key={selected} modelName={selected} />}
                  {tab === "setup" && (
                    <SetupPlan key={selected} modelName={selected} modelPhase={modelPhase}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} />
                  )}
                  {stage && (
                    <IssueTable key={selected + stage} modelName={selected} modelPhase={modelPhase}
                      heading={stage === "dev" ? "Dev Issues" : "Mass Production Issues"}
                      // a new issue starts in the model's phase when it belongs here, else in MP / no phase
                      newDefaults={{ phase: stage === "mp" ? "MP" : modelPhase === "MP" ? "" : modelPhase }}
                      entries={stageIssues} today={today} onChanged={(list) => onChanged(replaceSome(entries, stageIssues, list))} onError={fail} />
                  )}
                  {(tab === "dev" || tab === "review" || tab === "all") && (
                    <EntryList key={selected + tab} modelName={selected} kind={tab === "all" ? null : tab}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} onOpenIssues={(e) => setTab(issuesOfStage([e], "mp").length ? "issue" : "devissue")} />
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
