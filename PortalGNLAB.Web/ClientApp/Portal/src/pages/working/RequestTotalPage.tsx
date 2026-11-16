import { useEffect, useState } from "react";
import { useTr } from "./i18n";
import { METHODS, requestApi, type Method, type WorkingRequest } from "./request";
import { filtered, isCommon, modelCount, requestTotals, searched } from "./requestTotal";
import "./request.css";

/**
 * Working > Request Total: the saved PROCESSING REQUESTs added up, at a glance. Each JIG (name and drawing number both) is a
 * row with how many each model asked for; the JIGs common to models come first and can be shown alone.
 */
export default function RequestTotalPage(props: {
  onError: (e: unknown) => void;
  /** Already loaded (tests); loaded here when left out. */
  initial?: WorkingRequest[];
  /** Only the JIGs common to models at first (tests). */
  commonOnly?: boolean;
}) {
  const tr = useTr();
  const [list, setList] = useState<WorkingRequest[] | null>(props.initial ?? null);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [method, setMethod] = useState<Method | "">("");
  const [commonOnly, setCommonOnly] = useState(props.commonOnly ?? false);
  const [q, setQ] = useState("");

  useEffect(() => {
    if (props.initial) return;
    let live = true;
    requestApi.list().then((l) => { if (live) setList(l); }, (e) => { if (live) { setList([]); props.onError(e); } });
    return () => { live = false; };
  }, []);

  if (list === null) return <div className="muted small">{tr("Loading the requests…")}</div>;

  const t = requestTotals(filtered(list, { from, to, method }));
  const common = t.jigs.filter(isCommon).length;
  const shown = searched(commonOnly ? t.jigs.filter(isCommon) : t.jigs, q);
  const sum = (model: string) => shown.reduce((s, j) => s + (j.byModel[model] ?? 0), 0);

  return (
    <div className="rt-page">
      <div className="card rt-bar">
        <label className="rt-f">{tr("Request date")}
          <input type="date" className="inp sm" value={from} aria-label="From" onChange={(e) => setFrom(e.target.value)} />
          ~
          <input type="date" className="inp sm" value={to} aria-label="To" onChange={(e) => setTo(e.target.value)} />
        </label>
        <label className="rt-f">{tr("Processing method")}
          <select className="inp sm" value={method} aria-label="Processing method" onChange={(e) => setMethod(e.target.value as Method | "")}>
            <option value="">{tr("All")}</option>
            {METHODS.map(([k, label]) => <option key={k} value={k}>{label}</option>)}
          </select>
        </label>
        <label className="rt-f">
          <input type="checkbox" checked={commonOnly} onChange={(e) => setCommonOnly(e.target.checked)} />
          {tr("Only JIGs common to models")} ({common})
        </label>
        <input className="inp sm rt-q" value={q} placeholder="JIG name / drawing no" aria-label="Find a JIG" onChange={(e) => setQ(e.target.value)} />
        <span className="muted small rt-count">{t.requests} {tr("requests")} · {t.jigs.length} JIG · {t.models.length} {tr("models")}</span>
      </div>

      {!t.models.length ? <div className="card muted small rt-empty">{tr("No JIG requested yet.")}</div> : (
        <>
          <div className="rt-models">
            {t.models.map((m) => (
              <div key={m.model} className="card rt-model">
                <div className="rt-model-n">{m.model}</div>
                <div className="rt-model-t">{m.total.toLocaleString()}</div>
                <div className="muted small">{m.jigs} JIG · {m.requests} {tr("requests")}</div>
              </div>
            ))}
          </div>

          <div className="card rt-sheet">
            <table className="rt-t">
              <thead>
                <tr>
                  <th>No.</th><th className="rt-name">{tr("JIG name")}</th><th>Drawing No</th><th>{tr("Models")}</th>
                  {t.models.map((m) => <th key={m.model} className="rt-m">{m.model}</th>)}
                  <th className="rt-tot">Total</th><th>{tr("Requests")}</th>
                </tr>
              </thead>
              <tbody>
                {shown.map((j, n) => (
                  <tr key={j.key} className={isCommon(j) ? "rt-common" : ""}>
                    <td>{n + 1}</td>
                    <td className="rt-name">{j.name}{isCommon(j) && <span className="rq-tag none">{tr("common")}</span>}</td>
                    <td className="mono">{j.drawingNo || <span className="muted">—</span>}</td>
                    <td>{modelCount(j)}</td>
                    {t.models.map((m) => <td key={m.model} className="rt-m">{j.byModel[m.model] === undefined ? "" : j.byModel[m.model].toLocaleString()}</td>)}
                    <td className="rt-tot">{j.total.toLocaleString()}</td>
                    <td>{j.requests}</td>
                  </tr>
                ))}
                {!shown.length && <tr><td colSpan={t.models.length + 6} className="muted small">{tr("No JIG like this.")}</td></tr>}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan={4}>Total</td>
                  {t.models.map((m) => <td key={m.model} className="rt-m">{sum(m.model).toLocaleString()}</td>)}
                  <td className="rt-tot">{shown.reduce((s, j) => s + j.total, 0).toLocaleString()}</td><td />
                </tr>
              </tfoot>
            </table>
          </div>
        </>
      )}
    </div>
  );
}
