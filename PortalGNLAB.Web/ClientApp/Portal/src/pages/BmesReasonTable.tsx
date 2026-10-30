import { useEffect, useMemo, useRef, useState } from "react";
import { bmesApi, type Pair, type ReasonEdit, type ReasonRow } from "../bmesApi";
import "../bmes.css";
import StatusBar, { type Status } from "./bmes/StatusBar";
import { MISSING_KEY, OTHER_KEY, PAGE_SIZE, chips, displayRows, findPairsMessage, quickOptions, reasonOptions } from "./bmes/reasonLogic";

export default function BmesReasonTable() {
  const [rows, setRows] = useState<ReasonRow[]>([]);
  const [editId, setEditId] = useState(-1);
  const [newId, setNewId] = useState(-1); // row created by "+ Add row" and not saved yet
  const [draft, setDraft] = useState<ReasonEdit>({ processName: "", ngName: "", reason: "" });

  const [quickId, setQuickId] = useState(-1);
  const [quickValue, setQuickValue] = useState("");
  const [quickOther, setQuickOther] = useState(false);
  const quickSaving = useRef(false);
  const quickRef = useRef<HTMLSelectElement & HTMLInputElement>(null);

  const [search, setSearch] = useState("");
  const [chip, setChip] = useState("");
  const [limit, setLimit] = useState(PAGE_SIZE);

  const [modelInput, setModelInput] = useState("");
  const [activeModel, setActiveModel] = useState("");
  const [modelPairs, setModelPairs] = useState<Pair[] | null>(null);
  const [modelBusy, setModelBusy] = useState(false);

  const [busy, setBusy] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [status, setStatus] = useState<Status | null>(null);

  const error = (e: unknown) => setStatus({ text: (e as Error).message, error: true });
  const markDirty = (msg: string) => {
    setDirty(true);
    setStatus({ text: `${msg} Click Update To Server to sync.` });
  };

  async function reload(): Promise<void> {
    setRows((await bmesApi.reason.list()).rows);
  }

  useEffect(() => { reload().catch(error); }, []);
  useEffect(() => { quickRef.current?.focus(); }, [quickId, quickOther]);

  const options = useMemo(() => reasonOptions(rows), [rows]);
  const chipList = useMemo(() => chips(rows), [rows]);
  const display = useMemo(
    () => displayRows({ rows, chip, search, editId, modelPairs }),
    [rows, chip, search, editId, modelPairs],
  );
  const shown = display.slice(0, limit);

  async function run(work: () => Promise<void>) {
    try { await work(); } catch (e) { error(e); }
  }

  async function discardUnsavedNew(): Promise<void> {
    if (newId <= 0) return;
    await bmesApi.reason.remove(newId);
    if (editId === newId) setEditId(-1);
    setNewId(-1);
    await reload();
    setStatus({ text: "Unsaved new reason row discarded." });
  }

  function clearModel() {
    setModelPairs(null);
    setActiveModel("");
    setModelInput("");
    setLimit(PAGE_SIZE);
  }

  async function scanModel() {
    if (modelBusy) return;
    const pattern = modelInput.trim();
    if (!pattern) return clearModel();
    setModelBusy(true);
    setActiveModel(pattern);
    setModelPairs([]);
    setLimit(PAGE_SIZE);
    try {
      setModelPairs((await bmesApi.reason.modelPairs(pattern)).pairs);
    } catch (e) {
      setStatus({ text: `Model lookup failed: ${(e as Error).message}`, error: true });
    } finally {
      setModelBusy(false);
    }
  }

  const addBlank = () => run(async () => {
    if (modelPairs) clearModel();
    setChip("");
    setSearch("");
    await discardUnsavedNew();
    const blank = { processName: "", ngName: "", reason: "" };
    const { id } = await bmesApi.reason.add(blank);
    await reload();
    setQuickId(-1);
    setEditId(id);
    setNewId(id);
    setDraft(blank);
    markDirty("Reason row added locally.");
  });

  const addFromPair = (processName: string, ngName: string) => run(async () => {
    await discardUnsavedNew();
    const row = { processName, ngName, reason: "" };
    const { id } = await bmesApi.reason.add(row);
    await reload();
    setEditId(id);
    setDraft(row);
    markDirty("Reason row added locally.");
  });

  const startEdit = (r: ReasonRow | null) => run(async () => {
    if (!r) return;
    if (r.id !== newId) await discardUnsavedNew();
    setQuickId(-1);
    setEditId(r.id);
    setDraft({ processName: r.processName, ngName: r.ngName, reason: r.reason });
  });

  const saveEdit = () => run(async () => {
    await bmesApi.reason.update(editId, draft);
    setEditId(-1);
    setNewId(-1);
    setQuickId(-1);
    await reload();
    markDirty("Reason row saved locally.");
  });

  const cancelEdit = () => run(async () => {
    await discardUnsavedNew();
    setEditId(-1);
  });

  const startQuick = (r: ReasonRow | null) => run(async () => {
    if (!r || quickId === r.id) return;
    await discardUnsavedNew();
    setEditId(-1);
    setQuickId(r.id);
    setQuickValue((r.reason ?? "").trim());
    setQuickOther(false);
  });

  const closeQuick = () => { setQuickId(-1); setQuickOther(false); };

  async function saveQuick(id: number, value: string) {
    if (quickSaving.current || id <= 0 || quickId !== id) return;
    const row = rows.find((r) => r.id === id);
    if (!row) return setQuickId(-1);
    const next = value.trim();
    if ((row.reason ?? "") === next) return closeQuick();
    quickSaving.current = true;
    try {
      await bmesApi.reason.update(id, { processName: row.processName, ngName: row.ngName, reason: next });
      closeQuick();
      await reload();
      markDirty("Reason row saved locally.");
    } catch (e) {
      error(e);
    } finally {
      quickSaving.current = false;
    }
  }

  const remove = (id: number) => run(async () => {
    await bmesApi.reason.remove(id);
    if (editId === id) setEditId(-1);
    if (newId === id) setNewId(-1);
    if (quickId === id) setQuickId(-1);
    await reload();
    markDirty("Reason row deleted locally.");
  });

  const importFile = () => run(async () => {
    const { count } = await bmesApi.reason.importFile();
    setEditId(-1);
    await reload();
    if (count > 0) markDirty(`Imported ${count} reason rows locally.`);
    else setStatus({ text: "File not found or empty.", error: true });
  });

  async function findNewPairs() {
    if (busy) return;
    setBusy(true);
    setStatus(null);
    try {
      await discardUnsavedNew();
      const result = await bmesApi.reason.findPairs((m) => setStatus({ text: m }));
      setEditId(-1);
      await reload();
      const msg = findPairsMessage(result);
      if (msg.added) setDirty(true);
      setStatus({ text: msg.text, error: msg.error });
    } catch (e) {
      error(e);
    } finally {
      setBusy(false);
    }
  }

  async function updateToServer() {
    if (busy) return;
    setBusy(true);
    setStatus({ text: "Updating server..." });
    try {
      await discardUnsavedNew();
      const result = await bmesApi.reason.sync();
      if (result.succeeded) {
        setDirty(false);
        setStatus({ text: `Reason table updated to server (${result.rows} rows).` });
      } else {
        setStatus({ text: `Server update failed. Local DB is unchanged: ${result.message}`, error: true });
      }
    } catch (e) {
      setStatus({ text: `Server update failed: ${(e as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  }

  function quickCell(r: ReasonRow) {
    if (quickId === r.id) {
      const stop = (e: { stopPropagation: () => void }) => e.stopPropagation();
      const onKey = (e: React.KeyboardEvent) => {
        if (e.key === "Enter" && quickOther) void saveQuick(r.id, quickValue);
        else if (e.key === "Escape") closeQuick();
      };
      return quickOther ? (
        <input ref={quickRef} className="bm-input bm-edit" list="reason-options" placeholder="Type a reason" value={quickValue}
          onChange={(e) => setQuickValue(e.target.value)} onKeyDown={onKey} onBlur={() => void saveQuick(r.id, quickValue)} onClick={stop} />
      ) : (
        <select ref={quickRef} className="bm-input bm-edit" value={quickValue} onClick={stop} onKeyDown={onKey}
          onBlur={() => { if (!quickOther) setQuickId(-1); }}
          onChange={(e) => {
            if (e.target.value === OTHER_KEY) return setQuickOther(true);
            setQuickValue(e.target.value);
            void saveQuick(r.id, e.target.value);
          }}>
          <option value="">(empty)</option>
          {quickOptions(options, quickValue).map((o) => <option key={o} value={o}>{o}</option>)}
          <option value={OTHER_KEY}>Other...</option>
        </select>
      );
    }
    if (!(r.reason ?? "").trim()) return <span className="rs-hint">(empty - click to set)</span>;
    return <>{r.reason}</>;
  }

  const draftInput = (key: keyof ReasonEdit, list?: string) => (
    <input className="bm-input bm-edit" list={list} value={draft[key]} onChange={(e) => setDraft({ ...draft, [key]: e.target.value })} />
  );

  return (
    <div className="bm-page">
      <div className="bm-head">
        <div className="bm-title">
          <h1>Reason Table</h1>
          <span className="bm-sub">{rows.length} rows</span>
        </div>
        <div className="bm-toolbar">
          {dirty && <span className="bm-dirty" title="Local edits are not on the server until you click Update To Server.">Unsynced local changes</span>}
          <button className="btn sm primary" onClick={() => void findNewPairs()} disabled={busy} title="Fetch new (ProcessName, NgName) pairs from the BMES API, then from the cached daily NG data">Find new pairs</button>
          <button className={"btn sm" + (dirty ? " warn" : "")} onClick={() => void updateToServer()} disabled={busy} title="Send local Reason Table rows to the server DB">Update To Server</button>
          <button className="btn sm" onClick={() => void importFile()} disabled={busy} title="Import from reason.txt (local only; then Update To Server)">Import from file</button>
        </div>
      </div>

      <StatusBar status={status} busy={busy} onDismiss={() => setStatus(null)} />

      <div className="rs-filters">
        <div className="rs-chips">
          {chipList.map((c) => (
            <button key={c.key} className={"rs-chip" + (chip === c.key ? " on" : "") + (c.key === MISSING_KEY && c.count > 0 ? " warn" : "")}
              onClick={() => { setChip(c.key); setLimit(PAGE_SIZE); }}>{c.label} ({c.count})</button>
          ))}
        </div>
        <div className="rs-inputs">
          <input className="bm-input" style={{ width: 240 }} placeholder="Search ProcessName / NgName / Reason" value={search}
            onChange={(e) => { setSearch(e.target.value); setLimit(PAGE_SIZE); }} />
          <input className="bm-input" style={{ width: 170 }} placeholder="Model, Enter" value={modelInput}
            onChange={(e) => setModelInput(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") void scanModel(); }} />
          {modelPairs && <button className="btn sm" onClick={clearModel} title="Leave model view">Clear model</button>}
          <button className="btn sm" onClick={() => void addBlank()}>+ Add row</button>
        </div>
      </div>

      {modelPairs && (
        <div className="rs-model-note">
          Model <b>{activeModel}</b>: {modelBusy ? "scanning..." : `${modelPairs.length} pair(s)`} · pairs not in the Reason Table are marked red; use Add to create them.
        </div>
      )}

      <datalist id="reason-options">{options.map((o) => <option key={o} value={o} />)}</datalist>

      <div className="rs-table-wrap">
        <table className="bm-table">
          <thead>
            <tr>
              <th style={{ minWidth: 200 }}>ProcessName</th>
              <th style={{ minWidth: 200 }}>NgName</th>
              <th style={{ minWidth: 220 }}>Reason</th>
              <th style={{ width: 110 }}></th>
            </tr>
          </thead>
          <tbody>
            {shown.map((d, i) => {
              const r = d.row;
              if (r && r.id === editId) {
                return (
                  <tr key={r.id} className="editing">
                    <td>{draftInput("processName")}</td>
                    <td>{draftInput("ngName")}</td>
                    <td>{draftInput("reason", "reason-options")}</td>
                    <td className="actions">
                      <button className="btn sm ok" onClick={() => void saveEdit()}>Save</button>{" "}
                      <button className="btn sm" onClick={() => void cancelEdit()}>Cancel</button>
                    </td>
                  </tr>
                );
              }
              return (
                <tr key={r ? r.id : `pair-${i}-${d.processName}-${d.ngName}`} className={"row" + (r ? "" : " missing")} onDoubleClick={() => void startEdit(r)}>
                  <td>{d.processName}</td>
                  <td>{d.ngName}</td>
                  <td className={r ? "rs-reason-cell" : undefined} onClick={() => void startQuick(r)}>
                    {r ? quickCell(r) : <span className="rs-hint bad">(not in Reason Table)</span>}
                  </td>
                  <td className="actions">
                    {r ? (
                      <>
                        <button className="btn sm" onClick={() => void startEdit(r)} title="Edit">Edit</button>{" "}
                        <button className="btn sm danger" onClick={() => void remove(r.id)} title="Delete">Del</button>
                      </>
                    ) : (
                      <button className="btn sm" onClick={() => void addFromPair(d.processName, d.ngName)} title="Add to Reason Table">Add</button>
                    )}
                  </td>
                </tr>
              );
            })}
            {display.length === 0 && <tr><td colSpan={4} className="bm-empty">{modelBusy ? "Scanning..." : "No matching rows."}</td></tr>}
          </tbody>
        </table>
        {display.length > shown.length && (
          <div className="rs-more">
            Showing {shown.length} of {display.length}
            <button className="btn sm" onClick={() => setLimit(limit + PAGE_SIZE)}>Show more</button>
            <button className="btn sm" onClick={() => setLimit(Number.MAX_SAFE_INTEGER)}>Show all</button>
          </div>
        )}
      </div>
    </div>
  );
}
