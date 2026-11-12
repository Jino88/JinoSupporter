import { useEffect, useMemo, useState, type ClipboardEvent, type DragEvent } from "react";
import { workingApi } from "./api";
import {
  PLACE_TAB, PLACES, aiPictures, applyPlace, placeItems, placeRequest, readPlaceAnswer, type Place, type PlaceAnswer, type PlaceItem,
} from "./aiPlace";
import type { Entry, WorkingModel } from "./types";

type Asked = { answer: PlaceAnswer; items: PlaceItem[]; revision: number; pictures: File[] };

const dataUrl = (f: File) => new Promise<string>((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(String(r.result));
  r.onerror = () => reject(new Error(`${f.name || "A picture"} could not be read.`));
  r.readAsDataURL(f);
});

/** Small previews of pictures, each with its number (and a remove button while it can still be taken out). */
function Thumbs(props: { files: File[]; only?: number[]; onRemove?: (i: number) => void }) {
  const urls = useMemo(() => props.files.map((f) => URL.createObjectURL(f)), [props.files]);
  useEffect(() => () => urls.forEach((u) => URL.revokeObjectURL(u)), [urls]);
  const shown = props.only ?? props.files.map((_, i) => i);
  return (
    <span className="wk-ai-pics">
      {shown.filter((i) => i < urls.length).map((i) => (
        <span key={i} className="wk-ai-pic" title={props.files[i].name}>
          <img src={urls[i]} alt={`Picture ${i + 1}`} /><i>{i + 1}</i>
          {props.onRemove && <button type="button" aria-label={`Remove picture ${i + 1}`} onClick={() => props.onRemove!(i)}>×</button>}
        </span>
      ))}
    </span>
  );
}

/**
 * AI tab of a model: notes, mails or cells copied from Excel are typed or pasted, AI says where each piece belongs (Model
 * Info, Dev Plan, Dev Issue, Dev History, Reviews, MP Issues), and the items kept are saved there.
 */
export default function AiTab(props: {
  model: WorkingModel;
  entries: Entry[];
  today: string;
  onSaved: (model: WorkingModel) => void;
  onAdded: (entries: Entry[]) => void;
  onOpen: (tab: (typeof PLACE_TAB)[Place]) => void;
  onDirty: (dirty: boolean) => void;
  onError: (e: unknown) => void;
  /** An answer already there (tests). */
  initial?: PlaceAnswer;
}) {
  const { model } = props;
  const first = props.initial ? { answer: props.initial, items: placeItems(props.initial, model.info), revision: model.revision, pictures: [] } : null;
  const [text, setText] = useState("");
  const [pictures, setPictures] = useState<File[]>([]);
  const [busy, setBusy] = useState<AbortController | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [asked, setAsked] = useState<Asked | null>(first);
  const [picked, setPicked] = useState<Set<string>>(new Set(first?.items.map((i) => i.key)));
  const [done, setDone] = useState<[Place, number][]>([]);
  useEffect(() => props.onDirty(!!asked || !!text.trim() || pictures.length > 0), [asked, text, pictures]);

  /** Pictures pasted or dropped are kept to send; a text paste goes on into the box. */
  const addPictures = (files: File[]) => {
    const r = aiPictures(files, pictures.length);
    if (r.problem) setError(r.problem);
    else if (r.files.length) { setError(""); setPictures((cur) => [...cur, ...r.files]); }
    return r.files.length > 0 || !!r.problem;
  };
  const onPaste = (ev: ClipboardEvent) => { if (!busy && addPictures(Array.from(ev.clipboardData.files))) ev.preventDefault(); };
  const onDrop = (ev: DragEvent) => { if (ev.dataTransfer.files.length) { ev.preventDefault(); if (!busy) addPictures(Array.from(ev.dataTransfer.files)); } };

  const send = async () => {
    if ((!text.trim() && !pictures.length) || busy) return;
    const ctl = new AbortController();
    setBusy(ctl); setError(""); setAsked(null); setDone([]);
    try {
      const images = await Promise.all(pictures.map(async (f, i) => ({ name: f.name || `pasted-${i + 1}.png`, dataUrl: await dataUrl(f) })));
      const answer = readPlaceAnswer(await workingApi.placeAi(model.name, placeRequest(text, props.today, model.info, props.entries, images), ctl.signal));
      // a picture number AI made up is left out
      answer.entries = answer.entries.map((e) => ({ ...e, images: e.images.filter((n) => n < pictures.length) }));
      const items = placeItems(answer, model.info);
      setAsked({ answer, items, revision: model.revision, pictures });
      setPicked(new Set(items.map((i) => i.key)));
    } catch (e) {
      if (!ctl.signal.aborted) setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(null); }
  };

  const items = asked?.items ?? [];
  const isNew = (key?: string) => !!key && items.some((i) => i.key === key && !i.change);
  /** A task under a new event that is left out goes with it. */
  const off = (i: PlaceItem) => isNew(i.parent) && !picked.has(i.parent!);
  const kept = items.filter((i) => picked.has(i.key) && !off(i));
  const flip = (key: string) => setPicked((s) => { const n = new Set(s); if (n.has(key)) n.delete(key); else n.add(key); return n; });

  const apply = async () => {
    if (!asked || !kept.length || saving) return;
    if (asked.revision !== model.revision && !window.confirm("The model changed after AI answered. Put the items in anyway?")) return;
    const { info, entries, pictures: pics } = applyPlace(model.info, asked.answer, new Set(kept.map((i) => i.key)), props.today);
    setSaving(true); setError("");
    const added: Entry[] = [];
    try {
      if (info) props.onSaved(await workingApi.saveModel(model.name, info, model.revision));
      for (const [i, e] of entries.entries()) {
        const entry = await workingApi.addEntry(model.name, e);
        added.push(entry);
        const files = pics[i].map((n) => asked.pictures[n]).filter(Boolean);
        if (files.length) added[added.length - 1] = { ...entry, images: [...entry.images, ...await workingApi.addImages(model.name, entry.id, "", files)] };
      }
      setDone(PLACES.map((p): [Place, number] => [p, kept.filter((i) => i.place === p).length]).filter(([, n]) => n > 0));
      setAsked(null); setText(""); setPictures([]);
    } catch (e) {
      setError(`${info || added.length ? "Only part was saved: " : ""}${e instanceof Error ? e.message : String(e)}`);
    } finally {
      if (added.length) props.onAdded(added);
      setSaving(false);
    }
  };

  return (
    <section className="card pi-sec wk-ai">
      <div className="pi-sec-h">
        <h2>Place with AI</h2>
        <div className="pi-legend">{asked?.answer.model && <span>{asked.answer.model}</span>}</div>
      </div>
      <div className="pi-sec-b">
        <div className="muted small">
          Type or paste anything about {model.name} — meeting notes, a mail, a schedule, cells copied from Excel, or pictures
          (Ctrl+V a screenshot or a photo, or drop it here; up to 6). AI puts each
          piece where it belongs: Model Info, Dev Plan events, tasks and quantities (Material Plan / Making Plan), Dev Issues,
          Dev History, Reviews or Mass Production issues. You see every item first and keep the ones you want. Ctrl+Enter sends.
        </div>
        <textarea className="inp wk-ta wk-ai-text" rows={8} value={text} disabled={!!busy} aria-label="Text to place"
          placeholder={"e.g. DR build 11/20–11/25, jig ready by 11/21 (Park). Buzz noise found at 1 kHz, Major, fix by 11/30. SOP target 2027-03-01."}
          onPaste={onPaste} onDrop={onDrop} onDragOver={(e) => e.preventDefault()}
          onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void send(); } }} />
        {pictures.length > 0 && (
          <div className="wk-ai-sent">
            <span className="muted small">{pictures.length} picture{pictures.length > 1 ? "s" : ""} to send</span>
            <Thumbs files={pictures} onRemove={busy ? undefined : (i) => setPictures((cur) => cur.filter((_, n) => n !== i))} />
          </div>
        )}
        <div className="wk-ai-bar">
          <button className="btn sm primary" disabled={(!text.trim() && !pictures.length) || !!busy || saving} onClick={() => void send()}>{busy ? "AI is reading…" : "Place with AI"}</button>
          {busy && <button className="btn sm" onClick={() => busy.abort()}>Cancel</button>}
          {error && <span className="pi-dirty small" role="alert">{error}</span>}
          {done.length > 0 && (
            <span className="pi-msg small">
              Put in:{" "}
              {done.map(([p, n], i) => (
                <span key={p}>{i > 0 && ", "}<button className="wk-hm-link" onClick={() => props.onOpen(PLACE_TAB[p])}>{p}</button> {n}</span>
              ))}
            </span>
          )}
        </div>

        {asked && (
          <div className="wk-ai-answer">
            {asked.answer.note && <div className="wk-ai-note">{asked.answer.note}</div>}
            {!items.length ? <div className="muted small">AI found nothing to put in. Change the text and try again.</div> : (
              <>
                {PLACES.filter((p) => items.some((i) => i.place === p)).map((p) => (
                  <div key={p} className="wk-ai-place">
                    <h3>{p}<span className="cnt">{items.filter((i) => i.place === p && picked.has(i.key) && !off(i)).length}</span></h3>
                    {items.filter((i) => i.place === p).map((i) => (
                      <label key={i.key} className={"wk-ai-item" + (i.parent ? " sub" : "") + (off(i) ? " off" : "")}>
                        <input type="checkbox" checked={picked.has(i.key) && !off(i)} disabled={off(i)} onChange={() => flip(i.key)} />
                        <span className={"wk-ai-tag" + (i.remove ? " remove" : i.change ? " change" : "")}>{i.remove ? "Remove" : i.change ? "Change" : "New"}</span>
                        <span className="wk-ai-title">{i.title}</span>
                        {i.detail && <span className="wk-ai-detail">{i.detail}</span>}
                        {i.key.startsWith("entry.") && asked.answer.entries[Number(i.key.slice(6))].images.length > 0 && (
                          <Thumbs files={asked.pictures} only={asked.answer.entries[Number(i.key.slice(6))].images} />
                        )}
                      </label>
                    ))}
                  </div>
                ))}
                <div className="wk-ai-bar">
                  <button className="btn sm primary" disabled={!kept.length || saving} onClick={() => void apply()}>
                    {saving ? "Saving…" : `Put in ${kept.length} item${kept.length === 1 ? "" : "s"}`}
                  </button>
                  <button className="btn sm" disabled={saving} onClick={() => setPicked(new Set(items.map((i) => i.key)))}>All</button>
                  <button className="btn sm" disabled={saving} onClick={() => setPicked(new Set())}>None</button>
                  <button className="btn sm" disabled={saving} onClick={() => setAsked(null)}>Discard</button>
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </section>
  );
}
