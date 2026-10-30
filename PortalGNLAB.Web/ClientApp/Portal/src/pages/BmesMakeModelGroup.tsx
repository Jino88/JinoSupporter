import { useEffect, useMemo, useRef, useState } from "react";
import { bmesApi, type MaterialRow } from "../bmesApi";
import "../bmes.css";
import SubGroupNode from "./bmes/SubGroupNode";
import {
  PRODUCT_GROUPS, addGroup, addLineShift, addMid, addSubAtMidTop, addSubUnder, applyFocus, canMoveSub, clearGroups, dropSub,
  filterMaterials, focusFromQuery, fromApi, initialEditor, isCleared, midTotal, removeGroup, removeLineShift, removeMid, removeSub,
  renameGroup, renameMid, renameSub, resolveCategory, saveMessage, setActiveGroup, setActiveMid, setActiveSub, setProductGroup, toApi,
  type Editor,
} from "./bmes/modelGroupLogic";

interface Note { text: string; error: boolean }

export default function BmesMakeModelGroup() {
  const focus = useMemo(() => focusFromQuery(window.location.search), []);
  const [editor, setEditor] = useState<Editor>(initialEditor);
  const [allRows, setAllRows] = useState<MaterialRow[]>([]);
  const [mtypeMap, setMtypeMap] = useState<Record<string, string>>({});
  const [search, setSearch] = useState("");
  const [selected, setSelected] = useState("");
  const [selectedMtype, setSelectedMtype] = useState("");
  const [lineShifts, setLineShifts] = useState<string[]>([]);
  const [loadingMid, setLoadingMid] = useState(false);
  const [saving, setSaving] = useState(false);
  const [pulling, setPulling] = useState(false);
  const [note, setNote] = useState<Note | null>(null);
  const [draggingId, setDraggingId] = useState<string | null>(null);
  const selectToken = useRef(0);

  const { rows, missing } = useMemo(() => applyFocus(allRows, focus), [allRows, focus]);
  const filtered = useMemo(() => filterMaterials(rows, search, mtypeMap), [rows, search, mtypeMap]);
  const say = (text: string, error = false) => setNote({ text, error });

  async function selectRow(maktx: string, mtype: string) {
    const token = ++selectToken.current;
    setSelected(maktx);
    setSelectedMtype(mtype);
    setLoadingMid(true);
    setLineShifts([]);
    try {
      const { lineShifts: list } = await bmesApi.modelGroups.lineShifts(maktx);
      if (token === selectToken.current) setLineShifts(list);
    } catch (e) {
      if (token === selectToken.current) say((e as Error).message, true);
    } finally {
      if (token === selectToken.current) setLoadingMid(false);
    }
  }

  /** With the FCOST filter on, the first listed material is opened right away. */
  function selectFirstFocused(list: MaterialRow[]) {
    const first = applyFocus(list, focus).rows[0];
    if (focus.length > 0 && first) void selectRow(first.maktx, first.mtype);
  }

  async function pull(initial = false) {
    if (pulling) return;
    setPulling(true);
    setNote(null);
    try {
      const result = await bmesApi.modelGroups.pull();
      if (result.materialResult.succeeded) {
        setAllRows(result.view.rows);
        setMtypeMap(result.view.mtypeMap);
      }
      if (result.groupResult.succeeded) setEditor(fromApi(result.view.groups));

      const messages: string[] = [];
      messages.push(result.groupResult.succeeded
        ? result.groupResult.rows > 0
          ? `Loaded ${result.groupResult.rows} model group(s) from server DB.`
          : "Server DB has no model groups. Using local empty group."
        : `Server model group load failed. Using local DB: ${result.groupResult.message}`);
      messages.push(result.materialResult.succeeded
        ? `Loaded ${result.materialResult.rows.toLocaleString("en-US")} BMES material row(s) from server DB.`
        : `BMES Materials load failed: ${result.materialResult.message}`);
      say(messages.join(" "), !result.groupResult.succeeded || !result.materialResult.succeeded);

      if (result.materialResult.succeeded && (initial || !selected)) selectFirstFocused(result.view.rows);
    } catch (e) {
      say((e as Error).message, true);
    } finally {
      setPulling(false);
    }
  }

  async function reloadRows() {
    try {
      const view = await bmesApi.modelGroups.view();
      setAllRows(view.rows);
      setMtypeMap(view.mtypeMap);
    } catch (e) {
      say((e as Error).message, true);
    }
  }

  useEffect(() => {
    bmesApi.modelGroups.view()
      .then((view) => {
        setAllRows(view.rows);
        setMtypeMap(view.mtypeMap);
        if (view.groups.length > 0) {
          setEditor(fromApi(view.groups));
          selectFirstFocused(view.rows);
        } else {
          void pull(true); // nothing saved locally yet: take the server's groups, like the Blazor page
        }
      })
      .catch((e: Error) => say(e.message, true));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function save() {
    if (saving) return;
    setSaving(true);
    setNote(null);
    try {
      say("✓ " + saveMessage(await bmesApi.modelGroups.save(toApi(editor.groups))));
    } catch (e) {
      say(`Save failed: ${(e as Error).message}`, true);
    } finally {
      setSaving(false);
    }
  }

  const edit = (fn: (e: Editor) => Editor) => setEditor((e) => fn(e));

  function drop(targetId: string) {
    if (!draggingId) return;
    const sourceId = draggingId;
    setDraggingId(null);
    const result = dropSub(editor, sourceId, targetId);
    setEditor(result.editor);
    say(result.message, result.error);
  }

  const handlers = {
    activeSub: editor.activeSub,
    draggingId,
    canDrop: (id: string) => draggingId !== null && canMoveSub(editor, draggingId, id),
    onSelect: (id: string) => edit((e) => setActiveSub(e, id)),
    onAddNested: (id: string) => edit((e) => addSubUnder(e, id)),
    onRemove: (id: string) => edit((e) => removeSub(e, id)),
    onRename: (id: string, name: string) => edit((e) => renameSub(e, id, name)),
    onRemoveLineShift: (id: string, ls: string) => edit((e) => removeLineShift(e, id, ls)),
    onDragStart: (id: string) => { setDraggingId(id); setNote(null); },
    onDragEnd: () => setDraggingId(null),
    onDrop: drop,
  };

  return (
    <div className="bm-page">
      <div className="bm-head">
        <div className="bm-title">
          <h1>Model Group</h1>
          <span className="mg-badge">{rows.length} distinct</span>
        </div>
        <div className="bm-toolbar">
          <button className="btn sm" onClick={() => void reloadRows()}>↻ Refresh</button>
          <button className="btn sm" onClick={() => void pull()} disabled={pulling}>{pulling ? "Loading…" : "Load from server"}</button>
          <button className="btn sm primary" onClick={() => void save()} disabled={saving || pulling}>{saving ? "Saving…" : "💾 Save"}</button>
          <button className="btn sm danger" onClick={() => setEditor(clearGroups())} disabled={pulling || isCleared(editor.groups)}>🗑 Clear groups</button>
        </div>
      </div>

      {note && <div className={"mg-alert " + (note.error ? "err" : "ok")} role={note.error ? "alert" : "status"}>{note.text}</div>}

      {focus.length > 0 && (
        <div className="mg-alert">
          FCOST unmapped filter: showing {rows.length} / {focus.length} material(s).
          {missing.length > 0 && (
            <span style={{ color: "var(--err)", marginLeft: 4 }}>
              Not in BMES Materials: {missing.slice(0, 5).join(", ")}{missing.length > 5 ? ` … (+${missing.length - 5})` : ""}
            </span>
          )}
          <a style={{ marginLeft: 8 }} href="/app/bmes/make-model-group">Clear filter</a>
        </div>
      )}

      <div className="mg-cols">
        {/* Left: groups being built */}
        <div className="mg-card">
          <div className="mg-card-head">
            <span>Groups</span>
            <span className="mg-badge">{editor.groups.length}</span>
            <button className="btn sm primary mg-mini" style={{ marginLeft: "auto" }} title="Add new group (subsequent selections go here)"
              onClick={() => edit(addGroup)}>+ New group</button>
          </div>
          <div className="mg-scroll" style={{ padding: 8 }}>
            {editor.groups.map((g, gi) => {
              const active = gi === editor.activeGroup;
              return (
                <div key={g.id} className={"grp-box" + (active ? " active" : "")} onClick={() => edit((e) => setActiveGroup(e, gi))}>
                  <div className="grp-head">
                    {active && <span className="grp-dot" title="Selections go into this group">●</span>}
                    <input className="mg-name" value={g.name} placeholder="Group name"
                      onChange={(e) => edit((s) => renameGroup(s, gi, e.target.value))} onClick={(e) => e.stopPropagation()} />
                    <button className="btn sm mg-mini" title="Add model name" onClick={(e) => { e.stopPropagation(); edit((s) => addMid(s, gi)); }}>+ Model</button>
                    <button className="btn sm danger mg-mini" title="Remove group" onClick={(e) => { e.stopPropagation(); edit((s) => removeGroup(s, gi)); }}>✕</button>
                  </div>
                  <div className="pg-row" onClick={(e) => e.stopPropagation()}>
                    <span className="pg-label">Product</span>
                    {PRODUCT_GROUPS.map((pg) => (
                      <button key={pg} className={`pg-pill pg-${pg}` + (g.productGroup === pg ? " on" : "")}
                        onClick={() => edit((s) => setProductGroup(s, gi, pg))}>{pg}</button>
                    ))}
                  </div>
                  {g.midGroups.length === 0 ? (
                    <div className="mg-empty">(empty - click a Shift Group in the right panel or click "+ Model")</div>
                  ) : (
                    <div style={{ paddingBottom: 4 }}>
                      {g.midGroups.map((mid, mi) => {
                        const midActive = active && mi === editor.activeMid;
                        const hideHeader = mid.subGroups.length === 1 && !mid.subGroups[0].name && mid.subGroups[0].subGroups.length === 0;
                        return (
                          <div key={mid.id} className={"mid-grp" + (midActive ? " active" : "")}
                            onClick={(e) => { e.stopPropagation(); edit((s) => setActiveMid(s, gi, mi)); }}>
                            <div className="mid-head">
                              {midActive && <span className="grp-dot" title="Shift Groups go into the active model">●</span>}
                              <span className="tag-badge" title="Model Name">Model</span>
                              <input className="mg-name mid" value={mid.material} placeholder="Model"
                                onChange={(e) => edit((s) => renameMid(s, gi, mi, e.target.value))} onClick={(e) => e.stopPropagation()} />
                              <span className="bm-sub">{midTotal(mid)}</span>
                              <button className="btn sm mg-mini" title="Add sub group" onClick={(e) => { e.stopPropagation(); edit((s) => addSubAtMidTop(s, gi, mi)); }}>+ Sub Group</button>
                              <button className="btn sm danger mg-mini" title="Remove mid-group" onClick={(e) => { e.stopPropagation(); edit((s) => removeMid(s, gi, mi)); }}>✕</button>
                            </div>
                            <div className="sub-list">
                              {mid.subGroups.map((sub) => (
                                <SubGroupNode key={sub.id} node={sub} depth={0} hideHeader={hideHeader} {...handlers} />
                              ))}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Middle: distinct (Maktx, Mtype) */}
        <div className="mg-card">
          <div className="mg-card-head">BMES Materials — distinct (Maktx, Mtype)</div>
          <div style={{ padding: 8, borderBottom: "1px solid var(--line)" }}>
            <input className="bm-input" style={{ width: "100%" }} placeholder="Search Maktx / Mtype…" value={search} onChange={(e) => setSearch(e.target.value)} />
          </div>
          <div className="mg-scroll">
            {rows.length === 0 ? (
              <div className="mg-note">
                {focus.length > 0
                  ? "No BMES material rows matched the FCOST unmapped list."
                  : <>No data in BmesMaterials table.<br />Please run "Fetch all materials" in Setting first.</>}
              </div>
            ) : filtered.length === 0 ? (
              <div className="mg-note">No matches.</div>
            ) : (
              <table className="mg-table">
                <thead><tr><th style={{ width: 38 }}>#</th><th>Maktx</th><th style={{ width: 130 }}>Category</th></tr></thead>
                <tbody>
                  {filtered.map((r, i) => (
                    <tr key={`${r.maktx}\u0000${r.mtype}`} className={"click" + (selected === r.maktx && selectedMtype === r.mtype ? " selected" : "")}
                      onClick={() => { void selectRow(r.maktx, r.mtype); }}>
                      <td className="n">{i + 1}</td>
                      <td>{r.maktx}</td>
                      <td title={r.mtype}>{resolveCategory(r.mtype, mtypeMap)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>

        {/* Right: Shift Groups of the selected Maktx */}
        <div className="mg-card">
          <div className="mg-card-head">
            <span>Shift Groups</span>
            {selected && <span className="bm-sub">— {selected}</span>}
            {selected && <span className="mg-badge" style={{ marginLeft: "auto" }}>{lineShifts.length}</span>}
          </div>
          <div className="mg-scroll">
            {!selected ? (
              <div className="mg-note">Select a Maktx row in the middle panel.</div>
            ) : loadingMid ? (
              <div className="mg-note">Loading…</div>
            ) : lineShifts.length === 0 ? (
              <div className="mg-note">No Shift Group in DB for this MATERIALNAME.</div>
            ) : (
              <>
                <div className="mg-hint">Click a row to add it to the active mid-group</div>
                <table className="mg-table">
                  <thead><tr><th style={{ width: 38 }}>#</th><th>Shift Group</th></tr></thead>
                  <tbody>
                    {lineShifts.map((ls, i) => (
                      <tr key={ls} className="click" onClick={() => edit((e) => addLineShift(e, ls, selected))}>
                        <td className="n">{i + 1}</td>
                        <td>{ls}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
