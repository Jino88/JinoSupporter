// Display language of Working (English / Korean / Vietnamese, or the text as written). Texts are translated by AI on the
// server ahead of time (WorkingTranslationService); a text on screen is looked up there and shown as written until its
// translation is ready.
import { createContext, useCallback, useContext, useEffect, useRef, useState, type ReactNode } from "react";
import { workingApi } from "./api";

export type Lang = "" | "en" | "ko" | "vi";
export const LANGS: [Lang, string][] = [["", "Original"], ["en", "English"], ["ko", "Korean"], ["vi", "Vietnamese"]];
const KEY = "working.lang";

export const savedLang = (): Lang => {
  try {
    const v = localStorage.getItem(KEY) ?? "";
    return LANGS.some(([k]) => k === v) ? (v as Lang) : "";
  } catch { return ""; }
};
export const saveLang = (lang: Lang) => { try { localStorage.setItem(KEY, lang); } catch { /* private mode */ } };

/** A text in the language picked (as written when it has no translation yet). */
export type Tr = (text: string) => string;
export const TrContext = createContext<Tr>((t) => t);
export const useTr = () => useContext(TrContext);

/** The translation of a text from the ones known, or the text as written. */
export function lookUp(map: Record<string, string>, text: string): string {
  const k = text.trim();
  return k ? map[k] ?? text : text;
}

/**
 * Whether a text is already written in the language picked, so it is shown as written rather than through a translation
 * (which can reword it, e.g. a Korean note with a typo-like word changed). Korean: any Hangul in it.
 */
export const writtenIn = (lang: Lang, text: string) => lang === "ko" && /[\u3131-\u318e\uac00-\ud7a3]/.test(text);

/** How long to wait before asking again for texts the server is still translating, and how many times. */
const RETRY_MS = 15_000;
const RETRIES = 20;

/**
 * Gives the components under it `useTr()`. Texts asked for are gathered and sent to the server in one request a moment
 * later; the ones not translated yet are asked for again while the server translates them.
 */
export function TranslationProvider(props: { lang: Lang; children: ReactNode; onStatus?: (waiting: number) => void }) {
  const { lang } = props;
  const [map, setMap] = useState<Record<string, string>>({});
  const asked = useRef(new Set<string>());
  const pending = useRef(new Set<string>());
  const timer = useRef<number | null>(null);
  const tries = useRef(0);

  useEffect(() => {
    setMap({});
    asked.current = new Set();
    pending.current = new Set();
    tries.current = 0;
    props.onStatus?.(0);
    return () => { if (timer.current) window.clearTimeout(timer.current); timer.current = null; };
  }, [lang]);

  const send = useCallback(async () => {
    timer.current = null;
    const texts = [...pending.current];
    pending.current = new Set();
    if (!lang || !texts.length) return;
    try {
      const r = await workingApi.translations(lang, texts);
      setMap((cur) => ({ ...cur, ...r.translations }));
      const waiting = texts.filter((t) => r.translations[t] === undefined);
      props.onStatus?.(waiting.length);
      if (waiting.length && tries.current++ < RETRIES) {
        waiting.forEach((t) => pending.current.add(t));
        timer.current = window.setTimeout(() => void send(), RETRY_MS);
      }
    } catch {
      texts.forEach((t) => asked.current.delete(t)); // asked again when shown next
    }
  }, [lang]);

  const tr = useCallback<Tr>((text) => {
    if (!lang || !text) return text;
    const k = text.trim();
    if (!k || writtenIn(lang, k)) return text;
    if (map[k] !== undefined) return map[k];
    if (!asked.current.has(k)) {
      asked.current.add(k);
      pending.current.add(k);
      if (!timer.current) timer.current = window.setTimeout(() => void send(), 50);
    }
    return text;
  }, [lang, map, send]);

  return <TrContext.Provider value={tr}>{props.children}</TrContext.Provider>;
}
