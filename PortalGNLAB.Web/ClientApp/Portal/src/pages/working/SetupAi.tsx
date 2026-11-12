import { useEffect, useMemo, useState, type ClipboardEvent, type DragEvent } from "react";
import { aiPictures } from "./aiPlace";
import { setupApi, type SetupRow } from "./setup";
import { readSetupAnswer, setupAiItems, setupAiRequest, splitSetupAi, type SetupAiAnswer } from "./setupAiRows";

const dataUrl = (f: File) => new Promise<string>((resolve, reject) => {
  const r = new FileReader();
  r.onload = () => resolve(String(r.result));
  r.onerror = () => reject(new Error(`${f.name || "A picture"} could not be read.`));
  r.readAsDataURL(f);
});

/** Small previews of the pictures waiting to be sent, each with its number and a remove button. */
function Thumbs(props: { files: File[]; onRemove?: (i: number) => void }) {
  const urls = useMemo(() => props.files.map((f) => URL.createObjectURL(f)), [props.files]);
  useEffect(() => () => urls.forEach((u) => URL.revokeObjectURL(u)), [urls]);
  return (
    <span className="wk-ai-pics">
      {urls.map((u, i) => (
        <span key={i} className="wk-ai-pic" title={props.files[i].name}>
          <img src={u} alt={`Picture ${i + 1}`} /><i>{i + 1}</i>
          {props.onRemove && <button type="button" aria-label={`Remove picture ${i + 1}`} onClick={() => props.onRemove!(i)}>×</button>}
        </span>
      ))}
    </span>
  );
}

/**
 * Body of the Plan "AI input" popup: notes, mails, cells copied from Excel or pictures are typed or pasted, AI makes Setup
 * Plan rows of them, and the rows kept are added (New) or put over the row they change (Change).
 */
export default function SetupAi(props: { modelName: string; today: string; onDone: (savedCount: number) => void; onError: (e: unknown) => void }) {
  const { modelName } = props;
  const [rows, setRows] = useState<SetupRow[]>([]);
  const [text, setText] = useState("");
  const [pictures, setPictures] = useState<File[]>([]);
  const [busy, setBusy] = useState<AbortController | null>(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [answer, setAnswer] = useState<SetupAiAnswer | null>(null);
  const [asked, setAsked] = useState<SetupRow[]>([]);
  const [picked, setPicked] = useState<Set<string>>(new Set());

  useEffect(() => {
    let live = true;
    setupApi.rows(modelName).then((r) => { if (live) setRows(r); }, (e) => { if (live) props.onError(e); });
    return () => { live = false; };
  }, [modelName]);

  const items = useMemo(() => (answer ? setupAiItems(answer, asked) : []), [answer, asked]);
  const kept = items.filter((i) => picked.has(i.key));

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
    if ((!text.trim() && !pictures.length) || busy || saving) return;
    const ctl = new AbortController();
    setBusy(ctl); setError(""); setAnswer(null);
    try {
      const images = await Promise.all(pictures.map(async (f, i) => ({ name: f.name || `pasted-${i + 1}.png`, dataUrl: await dataUrl(f) })));
      const current = rows;
      const a = readSetupAnswer(await setupApi.ai(modelName, setupAiRequest(text, props.today, current, images), ctl.signal));
      setAsked(current); setAnswer(a);
      setPicked(new Set(setupAiItems(a, current).map((i) => i.key)));
    } catch (e) {
      if (!ctl.signal.aborted) setError(e instanceof Error ? e.message : String(e));
    } finally { setBusy(null); }
  };

  const flip = (key: string) => setPicked((s) => { const n = new Set(s); if (n.has(key)) n.delete(key); else n.add(key); return n; });

  const apply = async () => {
    if (!kept.length || saving) return;
    const { added, changed } = splitSetupAi(items, picked);
    setSaving(true); setError("");
    let saved = 0;
    try {
      if (added.length) saved += (await setupApi.import(modelName, added)).length;
      for (const c of changed) { await setupApi.save(modelName, c.row, c.fields); saved++; }
      props.onDone(saved);
    } catch (e) {
      setError(`${saved ? `Only ${saved} of ${kept.length} rows were saved: ` : ""}${e instanceof Error ? e.message : String(e)}`);
    } finally { setSaving(false); }
  };

  return (
    <div className="wk-ai">
      <div className="muted small">
        Type or paste anything about the setup of {modelName} — meeting notes, a mail, a schedule, cells copied from Excel, or
        pictures (Ctrl+V a screenshot or a photo, or drop it here; up to 6). AI makes Setup Plan rows of it. You see every row
        first and keep the ones you want. Ctrl+Enter sends.
      </div>
      <textarea className="inp wk-ta wk-ai-text" rows={8} value={text} disabled={!!busy} aria-label="Text to turn into setup rows"
        placeholder={"e.g. L1 Soldering: machine in on 11/20 (ready), setup 11/21–11/25 now setting up. L2 Packing not yet set up."}
        onPaste={onPaste} onDrop={onDrop} onDragOver={(e) => e.preventDefault()}
        onChange={(e) => setText(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void send(); } }} />
      {pictures.length > 0 && (
        <div className="wk-ai-sent">
          <span className="muted small">{pictures.length} picture{pictures.length > 1 ? "s" : ""} to send</span>
          <Thumbs files={pictures} onRemove={busy ? undefined : (i) => setPictures((cur) => cur.filter((_, n) => n !== i))} />
        </div>
      )}
      <div className="wk-ai-bar">
        <button className="btn sm primary" disabled={(!text.trim() && !pictures.length) || !!busy || saving} onClick={() => void send()}>{busy ? "AI is reading…" : "Make rows with AI"}</button>
        {busy && <button className="btn sm" onClick={() => busy.abort()}>Cancel</button>}
        {answer?.model && <span className="muted small">{answer.model}</span>}
        {error && <span className="pi-dirty small" role="alert">{error}</span>}
      </div>

      {answer && (
        <div className="wk-ai-answer">
          {answer.note && <div className="wk-ai-note">{answer.note}</div>}
          {!items.length ? <div className="muted small">AI found no row to add or change. Change the text and try again.</div> : (
            <>
              {items.map((i) => (
                <label key={i.key} className="wk-ai-item">
                  <input type="checkbox" checked={picked.has(i.key)} onChange={() => flip(i.key)} />
                  <span className={"wk-ai-tag" + (i.row ? " change" : "")}>{i.row ? "Change" : "New"}</span>
                  <span className="wk-ai-title">{i.title}</span>
                  {i.detail && <span className="wk-ai-detail">{i.detail}</span>}
                </label>
              ))}
              <div className="wk-ai-bar">
                <button className="btn sm primary" disabled={!kept.length || saving} onClick={() => void apply()}>
                  {saving ? "Saving…" : `Put in ${kept.length} row${kept.length === 1 ? "" : "s"}`}
                </button>
                <button className="btn sm" disabled={saving} onClick={() => setPicked(new Set(items.map((i) => i.key)))}>All</button>
                <button className="btn sm" disabled={saving} onClick={() => setPicked(new Set())}>None</button>
                <button className="btn sm" disabled={saving} onClick={() => setAnswer(null)}>Discard</button>
              </div>
            </>
          )}
        </div>
      )}
    </div>
  );
}
