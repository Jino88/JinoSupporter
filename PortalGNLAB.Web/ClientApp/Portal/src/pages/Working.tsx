import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { localToday } from "./productionInput/logic";
import AiPopup from "./working/AiPopup";
import AiTab from "./working/AiTab";
import type { AiTarget } from "./working/aiPlace";
import { workingApi } from "./working/api";
import DevPlan from "./working/DevPlan";
import EntryList from "./working/EntryList";
import Home from "./working/Home";
import { LANGS, TranslationProvider, saveLang, savedLang, useTr, type Lang } from "./working/i18n";
import { treeModels } from "./working/homeSummary";
import InfoTab from "./working/InfoTab";
import IssueTable from "./working/IssueTable";
import {
  buildTree, countByKind, filterTree, isOpenIssue, issuesOfStage, pathOf, replaceSome, sortEntries, type GroupNode, type IssueStage,
} from "./working/logic";
import SetupPlan from "./working/SetupPlan";
import { hideModels, treeApi, withAdded, type TreeSettings } from "./working/tree";
import TreeManager from "./working/TreeManager";
import { phaseName, type Entry, type GroupsResponse, type WorkingModel } from "./working/types";
import "./working/working.css";

/** The pages of the top menu: Home and the MP issues of every model. */
type HomeView = "home" | "mpissues";
type Tab = "info" | "plan" | "devissue" | "dev" | "review" | "all" | "setup" | "issue" | "ai";
/** Categories of a model page and the tabs under each (Model Info and AI have none). */
const CATEGORIES: { label: string; tabs: [Tab, string][] }[] = [
  { label: "Model Info", tabs: [["info", "Model Info"]] },
  { label: "Development", tabs: [["plan", "Dev Plan"], ["devissue", "Dev Issue"], ["dev", "Dev History"], ["review", "Reviews"], ["all", "Timeline"]] },
  { label: "Mass Production", tabs: [["setup", "Plan"], ["issue", "Issues"]] },
  { label: "AI", tabs: [["ai", "AI"]] },
];
const categoryOf = (tab: Tab) => CATEGORIES.find((c) => c.tabs.some(([k]) => k === tab))!;
const isTab = (v: string | null): v is Tab => CATEGORIES.some((c) => c.tabs.some(([k]) => k === v));
/** The tabs with AI input (a popup that makes what is pasted into the data of the tab) and what AI makes there. */
const AI_TARGET: Partial<Record<Tab, AiTarget>> = { plan: "plan", devissue: "devissue", dev: "dev", review: "review", all: "entries", issue: "issue" };
/** The issue list of a tab: Dev Issue (development) or Issues (mass production). */
const stageOf = (tab: Tab): IssueStage | null => (tab === "devissue" ? "dev" : tab === "issue" ? "mp" : null);

/** Keeps ?model=&tab= (or ?view=home) in the address so a model page or Home can be linked and survives reload. */
function writeQuery(home: boolean, model: string | null, tab: Tab, view: HomeView = "home") {
  const q = new URLSearchParams(window.location.search);
  if (model && !home) q.set("model", model); else q.delete("model");
  if (home) { q.set("view", view); q.delete("tab"); } else { q.delete("view"); q.set("tab", tab); }
  window.history.replaceState(null, "", `${window.location.pathname}?${q}`);
}

/** Working with the language its texts are shown in (kept in this browser). */
export default function Working() {
  const [lang, setLang] = useState<Lang>(savedLang());
  const [waiting, setWaiting] = useState(0);
  return (
    <TranslationProvider lang={lang} onStatus={setWaiting}>
      <WorkingPage lang={lang} waiting={waiting} onLang={(l) => { saveLang(l); setLang(l); }} />
    </TranslationProvider>
  );
}

function WorkingPage(props: { lang: Lang; waiting: number; onLang: (lang: Lang) => void }) {
  const tr = useTr();
  const query = new URLSearchParams(window.location.search);
  const [groups, setGroups] = useState<GroupsResponse | null>(null);
  const [treeSettings, setTreeSettings] = useState<TreeSettings>({ hidden: [], added: [] });
  const [managing, setManaging] = useState(false);
  const [filter, setFilter] = useState("");
  const [closed, setClosed] = useState<Set<string>>(new Set());
  // Home opens when asked for or when no model is in the address
  const [home, setHome] = useState(query.get("view") === "home" || query.get("view") === "mpissues" || !query.get("model"));
  const [homeView, setHomeView] = useState<HomeView>(query.get("view") === "mpissues" ? "mpissues" : "home");
  const [selected, setSelected] = useState<string | null>(home ? null : query.get("model"));
  const [tab, setTab] = useState<Tab>(isTab(query.get("tab")) ? (query.get("tab") as Tab) : "info");
  const [model, setModel] = useState<WorkingModel | null>(null);
  const [entries, setEntries] = useState<Entry[]>([]);
  const [infoDirty, setInfoDirty] = useState(false);
  // AI input popup of the tab, whether something in it would be lost, and a count that reopens the tab after AI put items in
  const [aiOpen, setAiOpen] = useState(false);
  const [aiDirty, setAiDirty] = useState(false);
  const [fresh, setFresh] = useState(0);
  const [error, setError] = useState("");
  const load = useRef(0);
  const today = localToday();

  // every model (Glossary + added here), then the ones not hidden in Manage
  const fullTree = useMemo(
    () => buildTree(withAdded(groups?.groups ?? [], treeSettings.added, treeSettings.groups, treeSettings.placed)), [groups, treeSettings]);
  const tree = useMemo(() => hideModels(fullTree, treeSettings.hidden), [fullTree, treeSettings.hidden]);
  const shown = useMemo(() => filterTree(tree, filter), [tree, filter]);
  const homeModels = useMemo(() => (groups ? treeModels(tree) : null), [groups, tree]);
  const fail = (e: unknown) => setError((e as Error).message);

  useEffect(() => {
    // without the tree settings (older server, or a failure) every Glossary model is shown
    const empty: TreeSettings = { hidden: [], added: [] };
    Promise.all([workingApi.groups(), treeApi.get().catch(() => empty)])
      .then(([g, t]) => {
        setGroups(g);
        setTreeSettings(t);
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
    writeQuery(home, selected, tab, homeView);
  }, [home, selected, tab, homeView]);

  useEffect(() => {
    setModel(null);
    setEntries([]);
    setInfoDirty(false);
    if (selected) reload(selected);
  }, [selected, reload]);

  const pick = (name: string) => {
    if (name === selected && !home) return;
    if (infoDirty && !window.confirm("Discard the unsaved changes?")) return;
    setHome(false);
    setSelected(name);
  };
  const goHome = (v: HomeView = "home") => {
    if (home && homeView === v) return;
    if (!home && infoDirty && !window.confirm("Discard the unsaved changes?")) return;
    setHome(true);
    setHomeView(v);
    setSelected(null);
  };
  /** A model opened from Home, on the tab asked for. */
  const openFromHome = (name: string, k: Tab) => {
    setTab(k);
    setHome(false);
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
    if (infoDirty && !window.confirm("Discard the unsaved changes?")) return;
    setInfoDirty(false);
    setTab(k);
  };
  const category = categoryOf(tab);
  const aiTarget = AI_TARGET[tab];
  const openAi = () => {
    if (infoDirty && !window.confirm("The unsaved changes on this tab are lost when AI puts its items in. Open AI input anyway?")) return;
    setAiDirty(false);
    setAiOpen(true);
  };
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
        <nav className="wk-menu" aria-label="Working">
          {([["home", "Home"], ["mpissues", "MP Issues"]] as [HomeView, string][]).map(([v, label]) => (
            <button key={v} className={home && homeView === v ? "on" : ""} aria-current={home && homeView === v ? "page" : undefined}
              onClick={() => goHome(v)}>{tr(label)}</button>
          ))}
          <span className="wk-lang" role="group" aria-label="Language">
            {props.lang && props.waiting > 0 && <span className="muted small" title="AI is translating the rest">Translating {props.waiting}…</span>}
            {LANGS.map(([k, label]) => (
              <button key={k || "orig"} className={props.lang === k ? "on" : ""} aria-pressed={props.lang === k} onClick={() => props.onLang(k)}>{label}</button>
            ))}
          </span>
        </nav>
      </header>
      {error && <div className="alert" role="alert">{error}</div>}

      <div className="wk-lay">
        <aside className="card wk-side">
          <div className="wk-side-h">
            <div className="wk-side-t">
              <h2>{tr("Model groups")}</h2>
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
              Source: Company Glossary{treeSettings.added.length ? ` + ${treeSettings.added.length} added` : ""}
              {treeSettings.hidden.length ? ` · ${treeSettings.hidden.length} hidden` : ""}
            </div>
          )}
        </aside>
        {managing && (
          <TreeManager tree={fullTree} settings={treeSettings} onChanged={setTreeSettings} onClose={() => setManaging(false)} onError={fail} />
        )}

        <section className="wk-main">
          {home ? (
            <Home models={homeModels} today={today} onOpen={openFromHome} view={homeView} />
          ) : !selected ? (
            <div className="card pi-empty">{tr("Pick a model on the left.")}</div>
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
                      {tr(c.label)}{n > 0 && <span className="cnt bad">{n}</span>}
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
                        {tr(label)}{n != null && n > 0 && <span className="cnt">{n}</span>}
                      </button>
                    );
                  })}
                  {aiTarget && model && (
                    <button className="btn sm wk-ai-open" title="Paste notes, a mail, Excel cells or pictures; AI makes them into the data of this tab"
                      onClick={openAi}>✦ AI input</button>
                  )}
                </div>
              )}

              {!model ? <div className="muted">Loading…</div> : (
                <>
                  {tab === "info" && (
                    <InfoTab key={selected} model={model} onDirty={setInfoDirty}
                      onSaved={(m) => { setModel(m); setInfoDirty(false); }} onError={fail} />
                  )}
                  {tab === "plan" && (
                    <DevPlan key={`${selected}:${fresh}`} model={model} onDirty={setInfoDirty}
                      onSaved={(m) => { setModel(m); setInfoDirty(false); }} onError={fail} />
                  )}
                  {tab === "setup" && (
                    <SetupPlan key={selected} modelName={selected} modelPhase={modelPhase}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} />
                  )}
                  {stage && (
                    <IssueTable key={`${selected}:${stage}:${fresh}`} modelName={selected} modelPhase={modelPhase}
                      heading={stage === "dev" ? "Dev Issues" : "Mass Production Issues"}
                      // a new issue starts in the model's phase when it belongs here, else in MP / no phase
                      newDefaults={{ phase: stage === "mp" ? "MP" : modelPhase === "MP" ? "" : modelPhase }}
                      entries={stageIssues} today={today} onChanged={(list) => onChanged(replaceSome(entries, stageIssues, list))} onError={fail} />
                  )}
                  {tab === "ai" && (
                    <AiTab key={selected} model={model} entries={entries} today={today} onDirty={setInfoDirty}
                      onSaved={setModel} onAdded={(list) => onChanged([...entries, ...list])}
                      onOpen={(k) => { setInfoDirty(false); setTab(k); }} onError={fail} />
                  )}
                  {(tab === "dev" || tab === "review" || tab === "all") && (
                    <EntryList key={`${selected}:${tab}:${fresh}`} modelName={selected} kind={tab === "all" ? null : tab}
                      entries={entries} today={today} onChanged={onChanged} onError={fail} onOpenIssues={(e) => setTab(issuesOfStage([e], "mp").length ? "issue" : "devissue")} />
                  )}
                  {aiOpen && aiTarget && (
                    <AiPopup title={`AI input · ${tr(category.tabs.find(([k]) => k === tab)![1])}`} dirty={aiDirty} onClose={() => setAiOpen(false)}>
                      <AiTab key={selected} model={model} entries={entries} today={today} target={aiTarget} onDirty={setAiDirty}
                        onSaved={setModel} onAdded={(list) => onChanged([...entries, ...list])}
                        onOpen={(k) => { setAiOpen(false); setInfoDirty(false); setTab(k); }} onError={fail}
                        onDone={() => { setAiOpen(false); setInfoDirty(false); setFresh((n) => n + 1); }} />
                    </AiPopup>
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
