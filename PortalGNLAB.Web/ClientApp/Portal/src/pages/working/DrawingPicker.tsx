import { useEffect, useState } from "react";
import { useTr } from "./i18n";
import { drawingApi, type DrawingMatch, type DrawingSuggestion, type RequestItem } from "./request";

/**
 * The drawing numbers the JIG/MC drawing DB has for a row's JIG / equipment name (and its model), best first; one is picked
 * for the row. The name can be changed and searched again.
 */
export default function DrawingPicker(props: {
  item: RequestItem; model: string; onPick: (m: DrawingMatch) => void; onClose: () => void;
  /** Already found (tests); else asked of the drawing DB. */
  initial?: DrawingSuggestion[];
}) {
  const tr = useTr();
  const [name, setName] = useState(props.item.name);
  const [asked, setAsked] = useState(props.initial ? props.item.name : "");
  const [list, setList] = useState<DrawingSuggestion[] | null>(props.initial ?? null);
  const [problem, setProblem] = useState("");

  const search = (q: string) => {
    if (!q.trim()) return;
    setAsked(q); setList(null); setProblem("");
  };
  useEffect(() => { if (!props.initial) search(props.item.name); }, []);
  useEffect(() => {
    if (!asked || props.initial) return;
    const ctl = new AbortController();
    drawingApi.suggest(asked, props.model, ctl.signal).then((r) => {
      if (!r.ok) { setProblem(r.error || "The drawing DB could not be opened."); setList([]); return; }
      setList(r.suggestions ?? []);
    }, (e) => { if (!ctl.signal.aborted) { setProblem(e instanceof Error ? e.message : String(e)); setList([]); } });
    return () => ctl.abort();
  }, [asked]);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") props.onClose(); };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  });

  return (
    <div className="modal-back rq-noprint" onClick={(e) => { if (e.target === e.currentTarget) props.onClose(); }}>
      <div className="modal rq-pick" role="dialog" aria-modal="true" aria-label="Pick a drawing no">
        <div className="wk-ai-pop-h">
          <h2>{tr("Pick a drawing no")} · {props.item.no}. {props.item.name}</h2>
          <button className="btn sm" aria-label="Close" onClick={props.onClose}>×</button>
        </div>
        <div className="rq-pick-q">
          <input className="inp sm" value={name} aria-label="JIG / equipment name to search" onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") search(name); }} />
          <button className="btn sm primary" disabled={!name.trim()} onClick={() => search(name)}>{tr("Search")}</button>
          {props.model && <span className="muted small">{tr("Model")}: {props.model}</span>}
        </div>
        {problem && <div className="pi-dirty small" role="alert">Drawing DB: {problem}</div>}
        {list === null ? <div className="muted small rq-pick-wait">{tr("Searching the drawing DB…")}</div>
          : !list.length ? (!problem && <div className="muted small rq-pick-wait">{tr("No drawing in the DB looks like this name. Change the name and search again.")}</div>)
          : (
            <table className="rq-pick-t">
              <thead><tr><th>Drawing no</th><th>Name</th><th>Model</th><th>Kind</th><th /></tr></thead>
              <tbody>
                {list.map((s) => (
                  <tr key={s.drawingNo}>
                    <td className="mono">{s.drawingNo}{s.latest === 0 && <span className="rq-tag rev">old</span>}</td>
                    <td>{s.name}{s.assy === 1 && <span className="rq-tag none">ASSY</span>}</td>
                    <td>{s.model}</td>
                    <td>{s.category}</td>
                    <td><button className="btn sm primary" onClick={() => props.onPick(s)}>{tr("Pick")}</button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
      </div>
    </div>
  );
}
