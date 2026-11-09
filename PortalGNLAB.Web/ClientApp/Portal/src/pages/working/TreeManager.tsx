import { useState, type ReactNode } from "react";
import type { GroupNode } from "./logic";
import { ADDED_GROUP, allModels, groupChoices, treeApi } from "./tree";
import type { BmesCheck, TreeSettings } from "./tree";

const SOURCE_LABEL: Record<string, string> = { bmes: "BMES", catalog: "BOM & Drawing list", group: "Make Model Group" };

/**
 * Manage the model tree: tick the models to show (the rest are hidden for everyone), add a model that is not in the
 * Company Glossary after checking its name in BMES, or remove an added one.
 */
export default function TreeManager(props: {
  /** The whole tree, added models included, nothing hidden. */
  tree: GroupNode[];
  settings: TreeSettings;
  onChanged: (s: TreeSettings) => void;
  onClose: () => void;
  onError: (e: unknown) => void;
}) {
  const { tree, settings } = props;
  const [hidden, setHidden] = useState(() => new Set(settings.hidden.map((h) => h.toLowerCase())));
  const [query, setQuery] = useState("");
  const [busy, setBusy] = useState(false);
  const [name, setName] = useState("");
  const [groupId, setGroupId] = useState("");
  const [check, setCheck] = useState<BmesCheck | null>(null);
  const [checking, setChecking] = useState(false);

  const models = allModels(tree);
  const added = new Set(settings.added.map((a) => a.name.toLowerCase()));
  const dirty = models.some((m) => hidden.has(m.toLowerCase()) !== settings.hidden.some((h) => h.toLowerCase() === m.toLowerCase()));
  const shownCount = models.filter((m) => !hidden.has(m.toLowerCase())).length;
  const q = query.trim().toLowerCase();

  const toggle = (m: string) => setHidden((s) => { const n = new Set(s); const k = m.toLowerCase(); if (n.has(k)) n.delete(k); else n.add(k); return n; });
  const setAll = (show: boolean, list: string[]) =>
    setHidden((s) => { const n = new Set(s); for (const m of list) { if (show) n.delete(m.toLowerCase()); else n.add(m.toLowerCase()); } return n; });

  const run = async (work: () => Promise<void>) => {
    setBusy(true);
    try { await work(); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const save = () => run(async () => props.onChanged(await treeApi.setHidden(models.filter((m) => hidden.has(m.toLowerCase())))));
  const runCheck = async () => {
    setChecking(true);
    try { setCheck(await treeApi.check(name.trim())); } catch (e) { props.onError(e); } finally { setChecking(false); }
  };
  const add = () => {
    const n = name.trim();
    if (check?.name !== n || !check.exact) {
      const why = check?.name === n ? "was not found in BMES" : "has not been checked in BMES";
      if (!window.confirm(`${n} ${why}. Add it anyway?`)) return;
    }
    void run(async () => {
      const a = await treeApi.add(n, groupId);
      props.onChanged({ ...settings, added: [...settings.added, a] });
      setName("");
      setCheck(null);
    });
  };
  const remove = (m: string) => {
    if (!window.confirm(`Remove the added model ${m} from the tree? Its Working data stays and comes back if it is added again.`)) return;
    void run(async () => {
      await treeApi.remove(m);
      props.onChanged({ hidden: settings.hidden.filter((h) => h.toLowerCase() !== m.toLowerCase()), added: settings.added.filter((a) => a.name.toLowerCase() !== m.toLowerCase()) });
    });
  };

  const renderNode = (n: GroupNode, depth: number): ReactNode => {
    const own = n.models.filter((m) => !q || m.toLowerCase().includes(q));
    const kids = n.children.map((c) => renderNode(c, depth + 1)).filter(Boolean);
    if (!own.length && !kids.length) return null;
    const all = allModels([n]);
    return (
      <div key={n.id} className={"wk-mg d" + Math.min(depth, 2)}>
        <div className="wk-mg-h">
          <b>{n.name}</b>
          <button className="wk-link" onClick={() => setAll(true, all)}>Show all</button>
          <button className="wk-link" onClick={() => setAll(false, all)}>Hide all</button>
        </div>
        {own.map((m) => (
          <label key={m} className="wk-mg-m">
            <input type="checkbox" checked={!hidden.has(m.toLowerCase())} onChange={() => toggle(m)} /> {m}
            {added.has(m.toLowerCase()) && (
              <>
                <span className="wk-mg-tag">added</span>
                <button className="wk-link danger" disabled={busy} onClick={(e) => { e.preventDefault(); remove(m); }}>Remove</button>
              </>
            )}
          </label>
        ))}
        {kids}
      </div>
    );
  };

  return (
    <div className="modal-back" onClick={props.onClose}>
      <div className="modal wk-mgr" role="dialog" aria-label="Manage models" onClick={(e) => e.stopPropagation()}>
        <div className="wk-mgr-h">
          <h2>Manage models</h2>
          <button className="btn sm" onClick={props.onClose}>Close</button>
        </div>

        <section className="wk-mgr-add">
          <h3>Add a model</h3>
          <div className="wk-mgr-row">
            <input className="inp sm" placeholder="Model name, e.g. TIU-C11-18-ZZ" value={name} aria-label="Model name"
              onChange={(e) => { setName(e.target.value); setCheck(null); }}
              onKeyDown={(e) => { if (e.key === "Enter" && name.trim()) void runCheck(); }} />
            <select className="inp sm" value={groupId} aria-label="Group" onChange={(e) => setGroupId(e.target.value)}>
              <option value="">Added (no group)</option>
              {groupChoices(tree).map(([id, label]) => <option key={id} value={id}>{label}</option>)}
            </select>
            <button className="btn sm" disabled={!name.trim() || checking} onClick={runCheck}>{checking ? "Checking…" : "Check in BMES"}</button>
            <button className="btn sm primary" disabled={!name.trim() || busy} onClick={add}>Add</button>
          </div>
          {check && (
            <div className={"wk-mgr-check " + (check.exact ? "ok" : "no")}>
              <div>
                {check.exact ? `✓ ${check.name} is a BMES model name.` : `✕ ${check.name} was not found as a BMES model name.`}
                {!check.bmesChecked && <span className="muted"> BMES itself was not searched (no BMES connection set up); only the local lists were.</span>}
                {check.bmesError && <span className="pi-dirty"> BMES search failed: {check.bmesError}</span>}
              </div>
              {check.candidates.length > 0 && (
                <div className="wk-mgr-cands">
                  <span className="muted small">Similar names:</span>
                  {check.candidates.map((c) => (
                    <button key={c.productName} className="wk-chip" title={`${c.productCode || "no code"} · ${SOURCE_LABEL[c.source] ?? c.source}`}
                      onClick={() => { setName(c.productName); setCheck(null); }}>{c.productName}</button>
                  ))}
                </div>
              )}
            </div>
          )}
        </section>

        <section>
          <div className="wk-mgr-row">
            <h3>Shown models <span className="muted small">{shownCount} of {models.length}</span></h3>
            <input className="inp sm" placeholder="Search…" value={query} aria-label="Search models" onChange={(e) => setQuery(e.target.value)} />
          </div>
          <div className="wk-mgr-list">{tree.filter((n) => n.id !== ADDED_GROUP || n.models.length).map((n) => renderNode(n, 0))}</div>
          <div className="modal-actions">
            {dirty && <span className="pi-dirty">Unsaved</span>}
            <button className="btn sm primary" disabled={!dirty || busy} onClick={save}>{busy ? "Saving…" : "Save shown models"}</button>
          </div>
        </section>
      </div>
    </div>
  );
}
