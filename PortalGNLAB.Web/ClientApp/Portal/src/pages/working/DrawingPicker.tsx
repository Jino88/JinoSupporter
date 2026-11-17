import { useEffect, useState } from "react";
import { useTr } from "./i18n";
import { drawingApi, drawingKey, type DrawingMatch, type DrawingSuggestion, type RequestItem } from "./request";

/**
 * The drawing numbers the JIG/MC drawing DB has for a row's JIG / equipment name (of its model only), best first; one is
 * picked for the row. A row with no name is searched by its drawing number (or the part typed), of any model. The text can be
 * changed and searched again. The other versions of the row's drawing number come first.
 */
export default function DrawingPicker(props: {
  item: RequestItem; model: string; onPick: (m: DrawingMatch) => void; onClose: () => void;
  /** Every version of the row's drawing number (-00, -01 …) the DB has. */
  versions?: DrawingMatch[];
  /** Already found (tests); else asked of the drawing DB. */
  initial?: DrawingSuggestion[];
}) {
  const tr = useTr();
  // the row's name, or its drawing number when it has no name
  const first = props.item.name.trim() ? props.item.name : props.item.drawingNo;
  const [name, setName] = useState(first);
  const [asked, setAsked] = useState(props.initial ? first : "");
  // nothing to search with neither (the versions alone are shown)
  const [list, setList] = useState<DrawingSuggestion[] | null>(props.initial ?? (first.trim() ? null : []));
  const others = (props.versions ?? []).some((v) => drawingKey(v.drawingNo) !== drawingKey(props.item.drawingNo));
  const [problem, setProblem] = useState("");

  const search = (q: string) => {
    if (!q.trim()) return;
    setAsked(q); setList(null); setProblem("");
  };
  useEffect(() => { if (!props.initial) search(first); }, []);
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
          <h2>{tr("Pick a drawing no")} · {props.item.no}. {props.item.name.trim() || props.item.drawingNo}</h2>
          <button className="btn sm" aria-label="Close" onClick={props.onClose}>×</button>
        </div>
        {others && (
          <>
            <div className="small rq-pick-h">{tr("Versions of this drawing no in the DB")}</div>
            <table className="rq-pick-t">
              <thead><tr><th>Drawing no</th><th>Name</th><th>Model</th><th>Kind</th><th /></tr></thead>
              <tbody>
                {props.versions!.map((v) => {
                  const now = drawingKey(v.drawingNo) === drawingKey(props.item.drawingNo);
                  return (
                    <tr key={v.drawingNo} className={now ? "rq-pick-now" : ""}>
                      <td className="mono">{v.drawingNo}{v.latest === 1 && <span className="rq-tag none">latest</span>}</td>
                      <td>{v.name}</td>
                      <td>{v.model}</td>
                      <td>{v.category}</td>
                      <td>{now ? <span className="muted small">{tr("in the row")}</span>
                        : <button className="btn sm primary" aria-label={`Pick ${v.drawingNo}`} onClick={() => props.onPick(v)}>{tr("Pick")}</button>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            <div className="small rq-pick-h">{tr("Or find it by name or drawing no")}</div>
          </>
        )}
        <div className="rq-pick-q">
          <input className="inp sm" value={name} aria-label="Name or drawing no to search" placeholder="Name or drawing no" onChange={(e) => setName(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") search(name); }} />
          <button className="btn sm primary" disabled={!name.trim()} onClick={() => search(name)}>{tr("Search")}</button>
          {props.model && <span className="muted small">{tr("Model")}: {props.model}</span>}
        </div>
        {problem && <div className="pi-dirty small" role="alert">Drawing DB: {problem}</div>}
        {list === null ? <div className="muted small rq-pick-wait">{tr("Searching the drawing DB…")}</div>
          : !list.length ? (!problem && <div className="muted small rq-pick-wait">{props.model
            ? tr("No drawing of this model in the DB looks like this name or drawing no. Change it and search again.")
            : tr("No drawing in the DB looks like this name or drawing no. Change it and search again.")}</div>)
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
