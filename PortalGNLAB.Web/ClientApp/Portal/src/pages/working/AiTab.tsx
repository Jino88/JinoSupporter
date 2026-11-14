import { useEffect, useMemo, useState, type ClipboardEvent, type DragEvent } from "react";
import { workingApi } from "./api";
import {
  PLACE_TAB, PLACES, aiPictures, applyPlace, placeItems, placeRequest, readPlaceAnswer, type AiTarget, type Place, type PlaceAnswer, type PlaceItem,
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

/** What AI makes of the text on each tab AI input is opened on, and an example of what to paste there. */
const TARGET_HINT: Record<AiTarget, [string, string]> = {
  plan: ["Dev Plan events with their tasks and quantities (Material Plan / Making Plan)", "e.g. DR build 11/20–11/25, jig ready by 11/21 (Park). F-PCB 720ea arrives 11/18."],
  devissue: ["Dev Issues (problem, severity, owner, due, cause)", "e.g. Buzz noise at 1 kHz on DR samples, Major, Kim, fix by 11/30."],
  dev: ["Dev History records (what was done or decided)", "e.g. 11/12 decided to change the damper glue to type B after the drop test."],
  review: ["Reviews (the result of a review or meeting: Pass, Conditional or Fail)", "e.g. DR review 11/14: conditional pass, buzz noise to be closed before DV."],
  issue: ["Mass Production issues", "e.g. Line 2 scratch on the cover since 11/10, 0.8% NG, Major, cause: jig wear."],
  entries: ["Dev History, Dev Issue, Review or Mass Production issue entries", "e.g. 11/14 DR review conditional pass. Buzz noise at 1 kHz, Major, fix by 11/30."],
  personal: ["your plans (a title, details, owner and due day each)", "e.g. DV report to the customer by 11/20. Book the drop test jig next week. Call Park about the F-PCB."],
};

/**
 * AI input of a model: notes, mails or cells copied from Excel are typed or pasted, AI says where each piece belongs (Model
 * Info, Dev Plan, Dev Issue, Dev History, Reviews, MP Issues), and the items kept are saved there. With a target (AI input
 * of a tab, shown in a popup) everything is made into the data of that tab only.
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
  /** The tab AI input was opened on: only its kind of data is made, and the box is shown without its card (popup). */
  target?: AiTarget;
  /** The items kept were all put in (popup: close it). */
  onDone?: () => void;
  /** An answer already there (tests). */
  initial?: PlaceAnswer;
}) {
  const { model, target } = props;
  // My Plans is kept under a model of the user's id, which is not a name to show
  const about = target === "personal" ? "your plans" : model.name;
  const first = props.initial ? { answer: props.initial, items: placeItems(props.initial, model.info, target, props.entries), revision: model.revision, pictures: [] } : null;
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
      const answer = readPlaceAnswer(await workingApi.placeAi(model.name, placeRequest(text, props.today, model.info, props.entries, images, target), ctl.signal));
      // a picture number AI made up is left out
      answer.entries = answer.entries.map((e) => ({ ...e, images: e.images.filter((n) => n < pictures.length) }));
      const items = placeItems(answer, model.info, target, props.entries);
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
      props.onDone?.();
    } catch (e) {
      setError(`${info || added.length ? "Only part was saved: " : ""}${e instanceof Error ? e.message : String(e)}`);
    } finally {
      if (added.length) props.onAdded(added);
      setSaving(false);
    }
  };

  const body = (
    <>
      <div className="muted small">
        {target ? (
          <>
            Type or paste anything for {about} — meeting notes, a mail, a schedule, cells copied from Excel, or pictures
            (Ctrl+V a screenshot or a photo, or drop it here; up to 6). AI makes it into {TARGET_HINT[target][0]}. You see
            every item first and keep the ones you want. Ctrl+Enter sends.
          </>
        ) : (
          <>
            Type or paste anything about {model.name} — meeting notes, a mail, a schedule, cells copied from Excel, or pictures
            (Ctrl+V a screenshot or a photo, or drop it here; up to 6). AI puts each
            piece where it belongs: Model Info, Dev Plan events, tasks and quantities (Material Plan / Making Plan), Dev Issues,
            Dev History, Reviews or Mass Production issues. You see every item first and keep the ones you want. Ctrl+Enter sends.
          </>
        )}
        {target && asked?.answer.model && <span> · {asked.answer.model}</span>}
      </div>
      <textarea className="inp wk-ta wk-ai-text" rows={8} value={text} disabled={!!busy} aria-label="Text to place" autoFocus={!!target}
        placeholder={target ? TARGET_HINT[target][1]
          : "e.g. DR build 11/20–11/25, jig ready by 11/21 (Park). Buzz noise found at 1 kHz, Major, fix by 11/30. SOP target 2027-03-01."}
        onPaste={onPaste} onDrop={onDrop} onDragOver={(e) => e.preventDefault()}
        onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void send(); } }} />
      {pictures.length > 0 && (
        <div className="wk-ai-sent">
          <span className="muted small">{pictures.length} picture{pictures.length > 1 ? "s" : ""} to send</span>
          <Thumbs files={pictures} onRemove={busy ? undefined : (i) => setPictures((cur) => cur.filter((_, n) => n !== i))} />
        </div>
      )}
      <div className="wk-ai-bar">
        <button className="btn sm primary" disabled={(!text.trim() && !pictures.length) || !!busy || saving} onClick={() => void send()}>{busy ? "AI is reading…" : target ? "Make with AI" : "Place with AI"}</button>
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
    </>
  );

  if (target) return <div className="wk-ai">{body}</div>;
  return (
    <section className="card pi-sec wk-ai">
      <div className="pi-sec-h">
        <h2>Place with AI</h2>
        <div className="pi-legend">{asked?.answer.model && <span>{asked.answer.model}</span>}</div>
      </div>
      <div className="pi-sec-b">{body}</div>
    </section>
  );
}
