import { useEffect, useState } from "react";
import type { PortalUser } from "../api";
import { localToday } from "./productionInput/logic";
import AiPopup from "./working/AiPopup";
import AiTab from "./working/AiTab";
import { workingApi } from "./working/api";
import { LANGS, TranslationProvider, saveLang, savedLang, useTr, type Lang } from "./working/i18n";
import IssueTable from "./working/IssueTable";
import { personalKey } from "./working/personal";
import type { Entry, WorkingModel } from "./working/types";
import "./working/working.css";

/** Working › Personal: the signed-in user's own plans, on the same register as a model's Issues; nobody else sees them. */
export default function WorkingPersonal(props: { user: PortalUser }) {
  const [lang, setLang] = useState<Lang>(savedLang());
  const [waiting, setWaiting] = useState(0);
  return (
    <TranslationProvider lang={lang} onStatus={setWaiting}>
      <PersonalPage user={props.user.name} lang={lang} waiting={waiting} onLang={(l) => { saveLang(l); setLang(l); }} />
    </TranslationProvider>
  );
}

export function PersonalPage(props: {
  user: string; lang: Lang; waiting: number; onLang: (lang: Lang) => void; initial?: Entry[]; initialModel?: WorkingModel;
}) {
  const tr = useTr();
  const key = personalKey(props.user);
  const [entries, setEntries] = useState<Entry[] | null>(props.initial ?? null);
  // the model the plans are kept under (AI input reads and saves through it), AI input open, something in it to lose, and a
  // count that shows the register again after AI put plans in
  const [model, setModel] = useState<WorkingModel | null>(props.initialModel ?? null);
  const [aiOpen, setAiOpen] = useState(false);
  const [aiDirty, setAiDirty] = useState(false);
  const [fresh, setFresh] = useState(0);
  const [error, setError] = useState("");
  const fail = (e: unknown) => setError((e as Error).message);
  const today = localToday();

  useEffect(() => {
    if (props.initial) return;
    Promise.all([workingApi.model(key), workingApi.entries(key)])
      .then(([m, list]) => { setModel(m); setEntries(list); setError(""); }).catch(fail);
  }, [key, props.initial]);

  const issues = (entries ?? []).filter((e) => e.kind === "issue");
  const onChanged = (list: Entry[]) => setEntries([...(entries ?? []).filter((e) => e.kind !== "issue"), ...list]);
  return (
    <div className="tools-page pi-page wk-page">
      <header className="page-head">
        <h1>Working · Personal</h1>
        <nav className="wk-menu" aria-label="Working">
          <span className="muted small">{tr("Only you can see these plans.")} ({props.user})</span>
          <span className="wk-lang" role="group" aria-label="Language">
            {props.lang && props.waiting > 0 && <span className="muted small" title="AI is translating the rest">Translating {props.waiting}…</span>}
            {LANGS.map(([k, label]) => (
              <button key={k || "orig"} className={props.lang === k ? "on" : ""} aria-pressed={props.lang === k} onClick={() => props.onLang(k)}>{label}</button>
            ))}
          </span>
        </nav>
      </header>
      {error && <div className="alert" role="alert">{error}</div>}
      {entries === null ? <div className="muted">Loading…</div> : (
        <>
          {model && (
            <div className="wk-subtabs">
              <button className="btn sm wk-ai-open" title="Paste notes, a mail, Excel cells or pictures; AI makes them into your plans"
                onClick={() => { setAiDirty(false); setAiOpen(true); }}>✦ AI input</button>
            </div>
          )}
          <IssueTable key={fresh} modelName={key} modelPhase="" heading="My Plans" newDefaults={{ phase: "" }} numberInOrder
            entries={issues} today={today} onChanged={onChanged} onError={fail} />
        </>
      )}
      {aiOpen && model && (
        <AiPopup title={`AI input · ${tr("My Plans")}`} dirty={aiDirty} onClose={() => setAiOpen(false)}>
          <AiTab model={model} entries={entries ?? []} today={today} target="personal" onDirty={setAiDirty}
            onSaved={setModel} onAdded={(list) => setEntries((cur) => [...(cur ?? []), ...list])}
            onOpen={() => setAiOpen(false)} onError={fail}
            onDone={() => { setAiOpen(false); setFresh((n) => n + 1); }} />
        </AiPopup>
      )}
    </div>
  );
}
