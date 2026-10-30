import { useEffect, useMemo, useState } from "react";
import { bmesApi, type RoutingEdit, type RoutingRow } from "../bmesApi";
import "../bmes.css";
import StatusBar, { type Status } from "./bmes/StatusBar";
import { filterModels, modelCounts, pickModel, selectedRows, selectedTypes } from "./bmes/routingLogic";

const blank = (modelName: string): RoutingEdit => ({ modelName, processCode: "", processName: "", processType: "" });

export default function BmesRoutingTable() {
  const [rows, setRows] = useState<RoutingRow[]>([]);
  const [selected, setSelected] = useState("");
  const [search, setSearch] = useState("");
  const [newModel, setNewModel] = useState("");
  const [typeFilter, setTypeFilter] = useState("");
  const [editId, setEditId] = useState(-1);
  const [newId, setNewId] = useState(-1); // row created by "+ Add" and not saved yet
  const [draft, setDraft] = useState<RoutingEdit>(blank(""));
  const [busy, setBusy] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [status, setStatus] = useState<Status | null>(null);
  const [loaded, setLoaded] = useState(false);

  const error = (e: unknown) => setStatus({ text: (e as Error).message, error: true });
  const markDirty = (msg: string) => {
    setDirty(true);
    setStatus({ text: `${msg} Click Update To Server to sync.` });
  };

  /** Reads the rows again. `model` is the model to keep (the first one when it has no rows). */
  async function reload(model: string): Promise<RoutingRow[]> {
    const { rows: fresh } = await bmesApi.routing.list();
    setRows(fresh);
    setSelected(pickModel(model, fresh));
    return fresh;
  }

  useEffect(() => {
    reload("").catch(error).finally(() => setLoaded(true));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const models = useMemo(() => modelCounts(rows), [rows]);
  const shownModels = useMemo(() => filterModels(models, rows, search), [models, rows, search]);
  const shownRows = useMemo(() => selectedRows(rows, selected, typeFilter), [rows, selected, typeFilter]);
  const types = useMemo(() => selectedTypes(rows, selected), [rows, selected]);

  /** An unsaved new row is removed again so an empty row never reaches Update To Server. */
  async function discardUnsavedNew(keepModel: string): Promise<void> {
    if (newId <= 0) return;
    await bmesApi.routing.remove(newId);
    if (editId === newId) setEditId(-1);
    setNewId(-1);
    await reload(keepModel);
    setStatus({ text: "Unsaved new process row discarded." });
  }

  async function run(work: () => Promise<void>) {
    try { await work(); } catch (e) { error(e); }
  }

  const selectModel = (name: string) => run(async () => {
    await discardUnsavedNew(name);
    setSelected(name);
    setTypeFilter("");
    setEditId(-1);
  });

  const addProcess = (model: string) => run(async () => {
    if (!model) return;
    await discardUnsavedNew(model);
    const { id } = await bmesApi.routing.add(blank(model));
    await reload(model);
    setSelected(model);
    setTypeFilter("");
    setEditId(id);
    setNewId(id);
    setDraft(blank(model));
    markDirty("Process row added locally.");
  });

  const addForNewModel = () => {
    const name = newModel.trim();
    if (!name) return;
    setNewModel("");
    setSearch("");
    void addProcess(name);
  };

  const startEdit = (r: RoutingRow) => run(async () => {
    if (r.id !== newId) await discardUnsavedNew(selected);
    setEditId(r.id);
    setDraft({ modelName: r.modelName, processCode: r.processCode, processName: r.processName, processType: r.processType });
  });

  const saveEdit = () => run(async () => {
    await bmesApi.routing.update(editId, draft);
    setEditId(-1);
    setNewId(-1);
    await reload(selected);
    markDirty("Process row saved locally.");
  });

  const cancelEdit = () => run(async () => {
    await discardUnsavedNew(selected);
    setEditId(-1);
  });

  const remove = (id: number) => run(async () => {
    await bmesApi.routing.remove(id);
    if (editId === id) setEditId(-1);
    if (newId === id) setNewId(-1);
    await reload(selected);
    markDirty("Process row deleted locally.");
  });

  const importFile = () => run(async () => {
    const { count } = await bmesApi.routing.importFile();
    setEditId(-1);
    await reload(selected);
    if (count > 0) markDirty(`Imported ${count} routing rows locally.`);
    else setStatus({ text: "File not found or empty.", error: true });
  });

  async function refreshFromBmes() {
    if (busy) return;
    setBusy(true);
    setStatus(null);
    let last = "";
    try {
      const { saved } = await bmesApi.routing.refresh((m) => { last = m; setStatus({ text: m }); });
      if (saved < 0) {
        setStatus({ text: last || "Refresh from BMES failed.", error: true });
      } else {
        setEditId(-1);
        await reload(selected);
        markDirty(`Refreshed from BMES: ${saved} row(s) received.`);
      }
    } catch (e) {
      setStatus({ text: `Refresh from BMES failed: ${(e as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  }

  async function updateToServer() {
    if (busy) return;
    setBusy(true);
    setStatus({ text: "Updating server..." });
    try {
      await discardUnsavedNew(selected);
      const result = await bmesApi.routing.sync();
      if (result.succeeded) {
        setDirty(false);
        setStatus({ text: `Routing table updated to server (${result.rows} rows).` });
      } else {
        setStatus({ text: `Server update failed. Local DB is unchanged: ${result.message}`, error: true });
      }
    } catch (e) {
      setStatus({ text: `Server update failed: ${(e as Error).message}`, error: true });
    } finally {
      setBusy(false);
    }
  }

  const field = (key: keyof RoutingEdit) => (
    <input className="bm-input bm-edit" value={draft[key]} onChange={(e) => setDraft({ ...draft, [key]: e.target.value })} />
  );

  return (
    <div className="bm-page">
      <div className="bm-head">
        <div className="bm-title">
          <h1>Routing Table</h1>
          <span className="bm-sub">{rows.length} rows · {models.length} models</span>
        </div>
        <div className="bm-toolbar">
          {dirty && <span className="bm-dirty" title="Local edits are not on the server until you click Update To Server.">Unsynced local changes</span>}
          <button className="btn sm primary" onClick={() => void refreshFromBmes()} disabled={busy} title="Add missing routing rows from BMES API (GN, last 3 months)">Refresh from BMES</button>
          <button className={"btn sm" + (dirty ? " warn" : "")} onClick={() => void updateToServer()} disabled={busy} title="Send local Routing Table rows to the server DB">Update To Server</button>
          <button className="btn sm" onClick={() => void importFile()} disabled={busy} title="Import from Routing.txt (local only; then Update To Server)">Import from file</button>
        </div>
      </div>

      <StatusBar status={status} busy={busy} onDismiss={() => setStatus(null)} />

      <div className="rt-grid">
        <div className="bm-panel">
          <div className="bm-panel-head">
            <input className="bm-input grow" placeholder="Search model, process code or name" value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <div className="rt-list">
            {shownModels.map((m) => (
              <div key={m.name} className={"rt-model" + (m.name === selected ? " on" : "")} onClick={() => void selectModel(m.name)}>
                <span className="name">{m.name}</span>
                <span className="count">{m.count}</span>
              </div>
            ))}
            {shownModels.length === 0 && <div className="bm-empty">{loaded ? "No matching models." : "Loading…"}</div>}
          </div>
          <div className="bm-panel-foot">
            <input className="bm-input grow" placeholder="New model name" value={newModel}
              onChange={(e) => setNewModel(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") addForNewModel(); }} />
            <button className="btn sm" disabled={!newModel.trim()} onClick={addForNewModel}>+ Add</button>
          </div>
        </div>

        <div className="bm-panel">
          <div className="bm-panel-head">
            <span className="rt-sel">{selected || "No model selected"}</span>
            <span className="bm-sub">{shownRows.length} process(es)</span>
            <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
              <select className="bm-input" value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)}>
                <option value="">All types</option>
                {types.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
              <button className="btn sm primary" disabled={!selected} onClick={() => void addProcess(selected)}>+ Add process</button>
            </div>
          </div>
          <div className="rt-table-wrap">
            <table className="bm-table">
              <thead>
                <tr>
                  <th style={{ minWidth: 130 }}>ProcessCode</th>
                  <th style={{ minWidth: 200 }}>ProcessName</th>
                  <th style={{ minWidth: 110 }}>ProcessType</th>
                  <th style={{ width: 90 }}></th>
                </tr>
              </thead>
              <tbody>
                {shownRows.map((r) => r.id === editId ? (
                  <tr key={r.id} className="editing">
                    <td>{field("processCode")}</td>
                    <td>{field("processName")}</td>
                    <td>{field("processType")}</td>
                    <td className="actions">
                      <button className="btn sm ok" onClick={() => void saveEdit()}>Save</button>{" "}
                      <button className="btn sm" onClick={() => void cancelEdit()}>Cancel</button>
                    </td>
                  </tr>
                ) : (
                  <tr key={r.id} className="row" onDoubleClick={() => void startEdit(r)}>
                    <td>{r.processCode}</td>
                    <td>{r.processName}</td>
                    <td>{r.processType}</td>
                    <td className="actions">
                      <button className="btn sm" onClick={() => void startEdit(r)} title="Edit">Edit</button>{" "}
                      <button className="btn sm danger" onClick={() => void remove(r.id)} title="Delete">Del</button>
                    </td>
                  </tr>
                ))}
                {shownRows.length === 0 && <tr><td colSpan={4} className="bm-empty">No process rows.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}
