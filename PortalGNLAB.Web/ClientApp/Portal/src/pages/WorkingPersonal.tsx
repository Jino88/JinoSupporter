import { useEffect, useState } from "react";
import type { PortalUser } from "../api";
import { localToday } from "./productionInput/logic";
import { workingApi } from "./working/api";
import { LANGS, TranslationProvider, saveLang, savedLang, useTr, type Lang } from "./working/i18n";
import IssueTable from "./working/IssueTable";
import { personalKey } from "./working/personal";
import type { Entry } from "./working/types";
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

export function PersonalPage(props: { user: string; lang: Lang; waiting: number; onLang: (lang: Lang) => void; initial?: Entry[] }) {
  const tr = useTr();
  const key = personalKey(props.user);
  const [entries, setEntries] = useState<Entry[] | null>(props.initial ?? null);
  const [error, setError] = useState("");
  const fail = (e: unknown) => setError((e as Error).message);
  const today = localToday();

  useEffect(() => {
    if (props.initial) return;
    workingApi.entries(key).then((list) => { setEntries(list); setError(""); }).catch(fail);
  }, [key, props.initial]);

  const issues = (entries ?? []).filter((e) => e.kind === "issue");
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
        <IssueTable modelName={key} modelPhase="" heading="My Plans" newDefaults={{ phase: "" }}
          entries={issues} today={today} onChanged={(list) => setEntries([...(entries ?? []).filter((e) => e.kind !== "issue"), ...list])} onError={fail} />
      )}
    </div>
  );
}
