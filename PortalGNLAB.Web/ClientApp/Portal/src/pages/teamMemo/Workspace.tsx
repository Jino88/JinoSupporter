import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { memoApi } from "../../teamMemoApi";
import { ago, canDropFolder, countsByFolder, descendants, filterMemos, scopedFolders, today } from "./logic";
import { MemoTranslator } from "./translator";
import { Avatar, LanguageBar } from "./shared";
import { CalendarView, Gantt } from "./Schedule";
import MemoDetail from "./MemoDetail";
import { statuses, statusLabel, type Folder, type Memo, type MemoView, type Summary } from "./types";
export default function Workspace({ board = false }: {
    board?: boolean;
}) {
    const [view, setView] = useState<MemoView>();
    const [memo, setMemo] = useState<Memo | null>(null);
    const [selected, setSelected] = useState<string | null>(new URLSearchParams(window.location.search).get("id"));
    const selectedRef = useRef(selected);
    selectedRef.current = selected;
    const [error, setError] = useState("");
    const [busy, setBusy] = useState(false);
    const [search, setSearch] = useState("");
    const [filter, setFilter] = useState("All");
    const [folder, setFolder] = useState<string | null>(null);
    const [collapsed, setCollapsed] = useState(new Set<string>());
    const [plan, setPlan] = useState(/\/memo\/(plan|meetings)/.test(window.location.pathname));
    const [calendar, setCalendar] = useState(plan);
    const [schedule, setSchedule] = useState(true);
    const [lang, setLanguage] = useState(() => sessionStorage.getItem("team-memo-language") ?? "");
    const [, redraw] = useState(0);
    const [translationVersion, setTranslationVersion] = useState(0);
    const translator = useMemo(() => new MemoTranslator({ cached: (lang, texts) => memoApi.json("/translations/cached", { lang, texts }), fetch: (lang, texts) => memoApi.json("/translations/fetch", { lang, texts }) }, () => redraw(n => n + 1)), []);
    const [form, setForm] = useState<{
        action: "new" | "rename" | "move";
        id: string;
        parentId: string | null;
        name: string;
    } | null>(null);
    const [drag, setDrag] = useState<{
        kind: "folder" | "memo";
        id: string;
    } | null>(null);
    const load = useCallback(async () => { const next = await memoApi.view(); setView(old => JSON.stringify(old) === JSON.stringify(next) ? old : next); const id = selectedRef.current; if (id) {
        try {
            const fresh = await memoApi.memo(id);
            if (selectedRef.current === id)
                setMemo(old => JSON.stringify(old) === JSON.stringify(fresh) ? old : fresh);
        }
        catch (e) {
            if (selectedRef.current === id) {
                setMemo(null);
                setError((e as Error).message);
            }
        }
    } }, []);
    useEffect(() => { let live = true; void load().catch(e => { if (live)
        setError(e.message); }); const timer = window.setInterval(() => { if (!document.hidden)
        void load().catch(e => { if (live)
            setError(e.message); }); }, 10000); return () => { live = false; clearInterval(timer); }; }, [load]);
    useEffect(() => { const onTranslation = () => setTranslationVersion(n => n + 1); window.addEventListener("memo-translations", onTranslation); return () => window.removeEventListener("memo-translations", onTranslation); }, []);
    useEffect(() => () => translator.dispose(), [translator]);
    const texts = JSON.stringify([...(view?.folders.flatMap(f => [f.name]) ?? []), ...(view?.memos.flatMap(m => [m.title, m.lastCommentBody]) ?? []), memo?.title, memo?.description, memo?.aiSummary, ...(memo?.tags.split(/[,;#\s]+/).filter(Boolean) ?? []), ...(memo?.comments.map(c => c.body) ?? []), ...(memo?.entries.flatMap(e => [e.title, e.body, e.aiSummary, ...e.comments.map(c => c.body)]) ?? [])]);
    useEffect(() => { void translator.refresh(lang, JSON.parse(texts)); }, [translator, lang, texts, translationVersion]);
    function setLang(value: string) { sessionStorage.setItem("team-memo-language", value); setLanguage(value); setTranslationVersion(n => n + 1); }
    async function run(action: () => Promise<void>) { setBusy(true); setError(""); try {
        await action();
        await load();
    }
    catch (e) {
        setError((e as Error).message);
    }
    finally {
        setBusy(false);
    } }
    function close() { selectedRef.current = null; setSelected(null); setMemo(null); const url = new URL(window.location.href); url.searchParams.delete("id"); window.history.replaceState(null, "", url); }
    async function open(id: string) { setError(""); selectedRef.current = id; setSelected(id); setMemo(null); const url = new URL(window.location.href); url.searchParams.set("id", id); window.history.replaceState(null, "", url); try {
        const fresh = await memoApi.memo(id);
        if (selectedRef.current === id)
            setMemo(fresh);
    }
    catch (e) {
        if (selectedRef.current === id)
            setError((e as Error).message);
    } }
    const display = useCallback((u: string) => view?.users.find(x => x.username.toLowerCase() === u.toLowerCase())?.displayName ?? u, [view]);
    const folders = useMemo(() => scopedFolders(view?.folders ?? [], plan), [view, plan]);
    const allowedIds = new Set(folders.map(f => f.id));
    const scoped = (view?.memos ?? []).filter(m => m.folderId ? allowedIds.has(m.folderId) : !plan);
    let visible = filterMemos(scoped, view?.folders ?? [], filter, search, view?.currentUser ?? "", display);
    if (folder) {
        const ids = descendants(folders, [folder]);
        visible = visible.filter(m => !!m.folderId && ids.has(m.folderId));
    }
    const counts = countsByFolder(visible, folders);
    const t = translator.t;
    function toggle(id: string) { setCollapsed(s => { const n = new Set(s); if (!n.delete(id))
        n.add(id); return n; }); }
    async function createMemo() { if (!view)
        return; await run(async () => { const created = await memoApi.json<{
        id: string;
    }>("/items", { folderId: folder ?? (plan ? folders.find(f => !f.parentId || !allowedIds.has(f.parentId))?.id : null), title: plan ? "New plan" : "New memo", kind: "Task" }); await open(created.id); }); }
    function drop(target: string | null) { if (!drag)
        return; const current = drag; setDrag(null); if (current.kind === "folder" && !canDropFolder(view?.folders ?? [], current.id, target)) {
        setError("A folder cannot be moved into itself or a descendant.");
        return;
    } void run(() => current.kind === "folder" ? memoApi.send("PUT", `/folders/${current.id}/move`, { parentId: target, name: "", plan: false }) : memoApi.send("POST", `/items/${current.id}/move`, { folderId: target })); }
    function memoNode(m: Summary) { return <button key={m.id} className={`memo-tree-memo ${selected === m.id ? "on" : ""}`} draggable onDragStart={e => { e.stopPropagation(); setDrag({ kind: "memo", id: m.id }); }} onDragEnd={() => setDrag(null)} onClick={() => void open(m.id)}><span>{m.isPinned ? "Pinned · " : ""}{t(m.title)}</span><small>{statusLabel(m.status)} · {m.entryCount} entries</small></button>; }
    function folderNodes(parent: string | null, seen = new Set<string>()): React.ReactNode { return folders.filter(f => (f.parentId && allowedIds.has(f.parentId) ? f.parentId : null) === parent && !seen.has(f.id)).sort((a, b) => a.ordinal - b.ordinal || a.name.localeCompare(b.name)).filter(f => !(search || filter !== "All") || (counts.get(f.id) ?? 0) > 0).map(f => { const next = new Set(seen); next.add(f.id); const folded = collapsed.has(f.id) && !search && filter === "All"; return <div className="memo-folder" key={f.id} onDragOver={e => { e.preventDefault(); e.stopPropagation(); }} onDrop={e => { e.preventDefault(); e.stopPropagation(); drop(f.id); }}><div className={`memo-folder-row ${folder === f.id ? "on" : ""}`} draggable onDragStart={e => { e.stopPropagation(); setDrag({ kind: "folder", id: f.id }); }} onDragEnd={() => setDrag(null)}><button aria-label="Expand or collapse folder" onClick={() => toggle(f.id)}>{folded ? "+" : "−"}</button><button className="memo-folder-name" onClick={() => setFolder(folder === f.id ? null : f.id)}>{t(f.name)} ({counts.get(f.id) ?? 0})</button><details><summary aria-label="Folder actions">Actions</summary><button onClick={() => setForm({ action: "new", id: "", parentId: f.id, name: "" })}>New subfolder</button><button onClick={() => setForm({ action: "rename", id: f.id, parentId: f.parentId, name: f.name })}>Rename</button><button onClick={() => setForm({ action: "move", id: f.id, parentId: f.parentId, name: f.name })}>Move</button><button onClick={() => void run(() => memoApi.send("PUT", `/folders/${f.id}/plan`, { parentId: null, name: "", plan: !plan }))}>{plan ? "Move to Work Memo" : "Move to Plan"}</button><button onClick={() => { if (window.confirm(`Delete folder '${f.path}'? Memos will be moved to its parent.`))
        void run(() => memoApi.send("DELETE", `/folders/${f.id}`)); }}>Delete</button></details></div>{!folded && <div className="memo-folder-children">{folderNodes(f.id, next)}{visible.filter(m => m.folderId === f.id).sort(sortMemos).map(memoNode)}</div>}</div>; }); }
    async function changePlan(value: boolean) { close(); setPlan(value); setFolder(null); setCalendar(value); if (value)
        await run(() => memoApi.send("POST", "/plans/initialize")); }
    return <div className="team-memo"><div className="memo-bar"><h1>{board ? "Memo Board" : plan ? "Plan" : "Work Memo"}</h1><a href={board ? "/app/memo" : "/app/memo/board"}>{board ? "Tree view" : "Board view"}</a>{!board && <><button className={!plan ? "on" : ""} onClick={() => void changePlan(false)}>Work Memo</button><button className={plan ? "on" : ""} onClick={() => void changePlan(true)}>Plan</button></>}<button disabled={busy || !view} onClick={() => void createMemo()}>New {plan ? "plan" : "memo"}</button><button disabled={busy} onClick={() => void run(load)}>Refresh</button></div>{error && <p role="alert">{error}<button onClick={() => setError("")}>Dismiss</button></p>}{!view ? <p>Loading memos...</p> : <><LanguageBar translator={translator} lang={lang} setLang={setLang} configured={view.translationConfigured}/><div className="memo-bar"><input aria-label="Search memos" placeholder="Search title, tag, folder or owner" value={search} onChange={e => setSearch(e.target.value)}/>{["All", "Mine", ...(!board ? ["Active"] : []), "Overdue"].map(f => <button className={filter === f ? "on" : ""} key={f} onClick={() => setFilter(f)}>{f}</button>)}<select aria-label="Filter folder" value={folder ?? ""} onChange={e => setFolder(e.target.value || null)}><option value="">All folders</option>{folders.map(f => <option key={f.id} value={f.id}>{f.path}</option>)}</select><button onClick={() => setCalendar(!calendar)}>{calendar ? "Hide calendar" : "Calendar"}</button>{board && <button onClick={() => setSchedule(!schedule)}>{schedule ? "Hide schedule" : "Schedule"}</button>}</div>{calendar && <CalendarView memos={visible} t={t} open={id => void open(id)}/>}<div className={board ? "memo-board-workspace" : "memo-tree-workspace"}>{!board && <aside className="memo-tree" onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); drop(null); }}><div className="memo-bar"><b>{visible.length} memos</b><button onClick={() => setForm({ action: "new", id: "", parentId: folder, name: "" })}>New folder</button></div>{form && <form className="memo-folder-form" onSubmit={e => { e.preventDefault(); void run(async () => { if (form.action === "new")
        await memoApi.send("POST", "/folders", { parentId: form.parentId, name: form.name, plan });
    else
        await memoApi.send("PUT", `/folders/${form.id}/${form.action}`, { parentId: form.parentId, name: form.name, plan }); setForm(null); }); }}><b>{form.action === "new" ? "New folder" : form.action === "move" ? "Move folder" : "Rename folder"}</b>{form.action !== "move" ? <input autoFocus aria-label="Folder name" value={form.name} onChange={e => setForm({ ...form, name: e.target.value })}/> : <select aria-label="Destination folder" value={form.parentId ?? ""} onChange={e => setForm({ ...form, parentId: e.target.value || null })}><option value="">Top level</option>{folders.filter(f => canDropFolder(view.folders, form.id, f.id)).map(f => <option key={f.id} value={f.id}>{f.path}</option>)}</select>}<button disabled={busy || form.action !== "move" && !form.name.trim()}>Save</button><button type="button" onClick={() => setForm(null)}>Cancel</button></form>}{folderNodes(null)}<h3>Unfiled</h3>{visible.filter(m => !m.folderId || !view.folders.some(f => f.id === m.folderId)).sort(sortMemos).map(memoNode)}</aside>}{board ? <div className="memo-board-main">{schedule && <Gantt memos={visible} t={t} open={id => void open(id)}/>}<div className="memo-columns">{statuses.map(status => <section className={`memo-column status-${status}`} key={status} onDragOver={e => e.preventDefault()} onDrop={e => { e.preventDefault(); if (drag?.kind === "memo") {
        const id = drag.id;
        setDrag(null);
        void run(() => memoApi.send("POST", `/items/${id}/status`, { status }));
    } }}><h2>{statusLabel(status)} ({visible.filter(m => m.status === status).length})</h2>{visible.filter(m => m.status === status).sort(sortMemos).map(m => <article className="memo-board-card" key={m.id} draggable onDragStart={() => setDrag({ kind: "memo", id: m.id })} onDragEnd={() => setDrag(null)}><button className="memo-card-title" onClick={() => void open(m.id)}>{m.isPinned ? "Pinned · " : ""}{t(m.title)}</button><p>{folders.find(f => f.id === m.folderId)?.path ?? "Unfiled"}</p><p>{m.startDate ?? "No start"} – {m.dueDate ?? "No due date"}</p><p>{m.tags}</p><div className="memo-bar">{m.participants.slice(0, 5).map(p => <Avatar key={p} name={p} display={display(p)}/>)}<small>{m.entryCount} entries · {m.commentCount} comments</small></div>{m.lastCommentBody && <blockquote>{m.lastCommentIsResult && <b>RESULT </b>}{display(m.lastCommentBy ?? "")}: {t(m.lastCommentBody)}</blockquote>}<small>Updated {ago(m.updatedAt)}</small></article>)}</section>)}</div></div> : <div>{memo ? <MemoDetail key={memo.id} memo={memo} view={view} t={t} changed={load} close={close} plan={plan}/> : <div className="memo-empty">{selected ? "Loading memo..." : "Select a memo, or create a new one."}</div>}</div>}</div>{board && memo && <div className="memo-drawer"><MemoDetail key={memo.id} memo={memo} view={view} t={t} changed={load} close={close}/></div>}</>}</div>;
}
function sortMemos(a: Summary, b: Summary) { return Number(b.isPinned) - Number(a.isPinned) || b.updatedAt.localeCompare(a.updatedAt); }
