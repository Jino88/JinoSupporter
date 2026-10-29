import { useEffect, useState, type ClipboardEvent } from "react";
import { memoApi } from "../../teamMemoApi";
import { ago, avatarColor, initials, markdownToHtml, size, stamp } from "./logic";
import type { AiSummary, Attachment, Memo, MemoView, Translate } from "./types";
import type { MemoTranslator } from "./translator";
export function Avatar({ name, display = name }: {
    name: string;
    display?: string;
}) { return <span className="memo-avatar" title={display} style={{ background: avatarColor(name) }}>{initials(display)}</span>; }
export function Attachments({ files, view, memoId, changed }: {
    files: Attachment[];
    view: MemoView;
    memoId: string;
    changed: () => Promise<void>;
}) {
    const [viewer, setViewer] = useState<Attachment | null>(null);
    const [error, setError] = useState("");
    async function remove(file: Attachment) { try {
        await memoApi.send("DELETE", `/items/${memoId}/attachments/${encodeURIComponent(file.id)}`);
        await changed();
    }
    catch (e) {
        setError((e as Error).message);
    } }
    return <><div className="memo-files">{files.map(file => <span key={file.id} className="memo-file">
  {file.isImage ? <button className="memo-image-button" onClick={() => setViewer(file)}><img loading="lazy" src={memoApi.attachment(file.id)} alt={file.fileName}/></button> : <a href={memoApi.attachment(file.id, true)}>{file.fileName}</a>}
  <small>{size(file.size)}</small>{(view.isAdmin || file.author.toLowerCase() === view.currentUser.toLowerCase()) && <button aria-label={`Delete ${file.fileName}`} onClick={() => void remove(file)}>Remove</button>}
 </span>)}</div>{error && <p role="alert">{error}</p>}{viewer && <div className="memo-viewer" role="dialog" aria-label="Image preview" onClick={() => setViewer(null)}><div onClick={e => e.stopPropagation()}><p>{viewer.fileName} · {size(viewer.size)} · {viewer.author} · {stamp(viewer.createdAt)}</p><a href={memoApi.attachment(viewer.id, true)}>Download</a> <a href={memoApi.attachment(viewer.id)} target="_blank" rel="noreferrer">Open original</a> <button onClick={() => setViewer(null)}>Close</button><img src={memoApi.attachment(viewer.id)} alt={viewer.fileName}/></div></div>}</>;
}
export const maxFileBytes = 20 * 1024 * 1024;
export function appendFiles(current: File[], incoming: File[]): File[] { if (current.length + incoming.length > 20)
    throw new Error("Select at most 20 files."); const large = incoming.find(f => f.size > maxFileBytes); if (large)
    throw new Error(`${large.name} is larger than 20 MB.`); return [...current, ...incoming]; }
export function pasteFiles(event: ClipboardEvent, add: (files: File[]) => void) { const files = Array.from(event.clipboardData.files); if (files.length) {
    event.preventDefault();
    add(files);
} }
function PendingFile({ file, remove }: {
    file: File;
    remove: () => void;
}) { const [url, setUrl] = useState(""); useEffect(() => { if (file.type.startsWith("image/") && file.size <= 4 * 1024 * 1024) {
    const value = URL.createObjectURL(file);
    setUrl(value);
    return () => URL.revokeObjectURL(value);
} }, [file]); return <span className="memo-file">{url && <img src={url} alt={file.name}/>}<span>{file.name} ({size(file.size)})</span><button onClick={remove}>Remove</button></span>; }
export function FilePicker({ files, setFiles, error }: {
    files: File[];
    setFiles: (files: File[]) => void;
    error: (text: string) => void;
}) { return <><input aria-label="Attach files" type="file" multiple onChange={e => { try {
    setFiles(appendFiles(files, Array.from(e.target.files ?? [])));
    e.target.value = "";
}
catch (ex) {
    error((ex as Error).message);
} }}/><div className="memo-files">{files.map((f, i) => <PendingFile key={i} file={f} remove={() => setFiles(files.filter((_, j) => i !== j))}/>)}</div></>; }
export function AiSummaryPanel({ summary, memo, entryId = "", view, t, changed }: {
    summary: AiSummary;
    memo: Memo;
    entryId?: string;
    view: MemoView;
    t: Translate;
    changed: () => Promise<void>;
}) {
    const [busy, setBusy] = useState(false);
    const [running, setRunning] = useState<string | null>(null);
    const [hidden, setHidden] = useState(false);
    const [error, setError] = useState("");
    useEffect(() => { let live = true; const poll = () => memoApi.get<{
        runningBy: string | null;
        entries: Record<string, string | null>;
    }>(`/items/${memo.id}/ai-status`).then(s => { if (live)
        setRunning(entryId ? s.entries?.[entryId] : s.runningBy); }).catch(() => { }); void poll(); const timer = window.setInterval(() => void poll(), 5000); return () => { live = false; clearInterval(timer); }; }, [memo.id, entryId]);
    async function generate() { setBusy(true); setError(""); try {
        await memoApi.send("POST", `/items/${memo.id}/summary`, { id: entryId });
        setHidden(false);
        await changed();
    }
    catch (e) {
        setError((e as Error).message);
    }
    finally {
        setBusy(false);
    } }
    return <section className="memo-ai"><div className="memo-bar"><b>AI Summary</b>{summary.isAiSummaryStale && <span className="memo-late">Outdated</span>}<small>{summary.aiSummaryBy} · {ago(summary.aiSummaryAt)} · {summary.aiSummaryModel}</small><button disabled={busy || !!running || !view.aiConfigured} onClick={() => void generate()}>{busy ? "Summarizing..." : running ? `${running} is summarizing...` : summary.aiSummary ? "Regenerate" : "Generate"}</button>{summary.aiSummary && <button onClick={() => setHidden(!hidden)}>{hidden ? "Show" : "Hide"}</button>}</div>{!view.aiConfigured && <p>Configure the Codex API key in Settings to use AI Summary.</p>}{error && <p role="alert">{error}</p>}{summary.aiSummary && !hidden && <div className="memo-markdown" dangerouslySetInnerHTML={{ __html: markdownToHtml(t(summary.aiSummary)) }}/>}</section>;
}
export function LanguageBar({ translator, lang, setLang, configured }: {
    translator: MemoTranslator;
    lang: string;
    setLang: (lang: string) => void;
    configured: boolean;
}) {
    const [status, setStatus] = useState<{
        running: boolean;
        done: number;
        total: number;
        message: string | null;
    }>();
    const [busy, setBusy] = useState(false);
    const [error, setError] = useState("");
    useEffect(() => { let live = true; const poll = () => memoApi.get<{
        running: boolean;
        done: number;
        total: number;
        message: string | null;
    }>("/translations/status").then(s => { if (live)
        setStatus(s); }).catch(() => { }); void poll(); const timer = window.setInterval(() => void poll(), 1500); return () => { live = false; clearInterval(timer); }; }, []);
    async function all() { setBusy(true); setError(""); try {
        await memoApi.send("POST", "/translations/all");
        setLang(lang);
        window.dispatchEvent(new Event("memo-translations"));
    }
    catch (e) {
        setError((e as Error).message);
    }
    finally {
        setBusy(false);
    } }
    return <div className="memo-language">{[["", "Original"], ["en", "English"], ["ko", "Korean"], ["vi", "Vietnamese"]].map(([code, label]) => <button key={code} className={code === lang ? "on" : ""} onClick={() => setLang(code)}>{label}</button>)}<button disabled={!configured || busy || status?.running} onClick={() => void all()}>Translate all</button><small role="status">{status?.running ? `All memos ${status.done}/${status.total}` : translator.busy ? `Translating ${translator.pending}...` : status?.message}</small>{(translator.error || error) && <span role="alert">{translator.error || error}</span>}</div>;
}
