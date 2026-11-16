import { useState } from "react";
import { localToday } from "./working/dates";
import { LANGS, TranslationProvider, saveLang, savedLang, useTr, type Lang } from "./working/i18n";
import RequestPage from "./working/RequestPage";
import type { WorkingRequest as Saved } from "./working/request";
import "./working/working.css";

/** Working › Request: the PROCESSING REQUEST forms (JIG/MC) AI fills from a pasted list or pictures. */
export default function WorkingRequest() {
  const [lang, setLang] = useState<Lang>(savedLang());
  const [waiting, setWaiting] = useState(0);
  return (
    <TranslationProvider lang={lang} onStatus={setWaiting}>
      <RequestShell lang={lang} waiting={waiting} onLang={(l) => { saveLang(l); setLang(l); }} />
    </TranslationProvider>
  );
}

export function RequestShell(props: { lang: Lang; waiting: number; onLang: (lang: Lang) => void; initial?: Saved[] }) {
  const tr = useTr();
  const [error, setError] = useState("");
  return (
    <div className="tools-page pi-page wk-page">
      <header className="page-head">
        <h1>Working · Request</h1>
        <nav className="wk-menu" aria-label="Working">
          <span className="muted small">{tr("JIG/MC processing requests")}</span>
          <span className="wk-lang" role="group" aria-label="Language">
            {props.lang && props.waiting > 0 && <span className="muted small" title="AI is translating the rest">Translating {props.waiting}…</span>}
            {LANGS.map(([k, label]) => (
              <button key={k || "orig"} className={props.lang === k ? "on" : ""} aria-pressed={props.lang === k} onClick={() => props.onLang(k)}>{label}</button>
            ))}
          </span>
        </nav>
      </header>
      {error && <div className="alert" role="alert">{error}</div>}
      <RequestPage today={localToday()} onError={(e) => setError((e as Error).message)} initial={props.initial} />
    </div>
  );
}
