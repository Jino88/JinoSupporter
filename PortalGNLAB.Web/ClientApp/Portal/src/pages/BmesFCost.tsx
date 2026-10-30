import { Fragment, useEffect, useMemo, useState, type CSSProperties, type ReactNode } from "react";
import { bmesReportApi, type FCostRawStatus, type FCostSetup, type FCostView } from "../bmesReportApi";
import { downloadBlob } from "../toolsApi";
import "../bmesReport.css";
import { csvText, isoDate, n0, stamp } from "./bmesReport/format";
import {
  buildCsvRows, buildHierarchy, colBoundaryClass, computeTotals, defaultStartDate, formatNum, formatRate, isRowVisible,
  levelLabel, metricCount, periodColumnCount, rateOf, unmappedModelGroupUrl, valueAt, visibleColumnIndexes,
  visiblePeriodColumnCount, COLLAPSE_LIMIT, type Expanded, type HierarchyRow, type Metrics,
} from "./bmesReport/fcostLogic";

const zero = (v: number) => (v === 0 ? "zero" : "");
const KINDS = ["Day", "Week", "Month"] as const;

interface MetricRowsProps {
  metrics: Metrics;
  rowClass: string;
  label: ReactNode;
  labelClass?: string;
  labelStyle?: CSSProperties;
  onLabelClick?: () => void;
  columns: number[];
  cellClass: (ci: number) => string;
  input: (ci: number, i: number) => number;
  cost: (ci: number, i: number) => number;
}

/** INPUT / F-COST / RATE % rows of one entity; the label cell spans the visible rows. */
function MetricRows({ metrics, rowClass, label, labelClass, labelStyle, onLabelClick, columns, cellClass, input, cost }: MetricRowsProps) {
  const span = metricCount(metrics);
  const kinds = [
    metrics.input && { key: "input", badge: <span className="ztype-badge zt-input">INPUT</span>, value: input, fmt: formatNum },
    metrics.cost && { key: "cost", badge: <span className="ztype-badge zt-cost">F-COST</span>, value: cost, fmt: formatNum },
    metrics.rate && {
      key: "rate", badge: <span className="ztype-badge zt-rate">RATE %</span>,
      value: (ci: number, i: number) => rateOf(input(ci, i), cost(ci, i)), fmt: formatRate,
    },
  ].filter(Boolean) as { key: string; badge: ReactNode; value: (ci: number, i: number) => number; fmt: (v: number) => string }[];

  return (
    <>
      {kinds.map((k, n) => (
        <tr key={k.key} className={`${rowClass} fcost-row-${k.key}`}>
          {n === 0 && (
            <td className={"label-td " + (labelClass ?? "")} rowSpan={span} style={labelStyle} onClick={onLabelClick}>{label}</td>
          )}
          <td className="label-td">{k.badge}</td>
          {columns.map((ci, i) => {
            const v = k.value(ci, i);
            return <td key={ci} className={`${zero(v)} ${cellClass(ci)}`}>{k.fmt(v)}</td>;
          })}
        </tr>
      ))}
    </>
  );
}

function RawStatusLine({ status }: { status: FCostRawStatus | null }) {
  if (!status?.exists) return <div className="muted">No RAW DB yet. Generate Report will create it.</div>;
  return (
    <div className="muted">
      Pulls {n0(status.successCount)} OK / {n0(status.failedCount)} failed, rows {n0(status.totalRows)}, range {status.firstDate} - {status.lastDate}
    </div>
  );
}

/**
 * F-Cost (port of BmesFCostPage.razor as reached at /bmes/f-cost): builds the report from the local
 * RAW cache through BmesFCostReportCalculationService, then the Model / Sub Group summary (expandable
 * Mid → Sub1 → Sub2) and the per-material table, with column folding and the CSV export. The Razor
 * component stays for the combined report HTML; this is a separate view over the same service.
 */
export default function BmesFCost() {
  const [setup, setSetup] = useState<FCostSetup | null>(null);
  const [start, setStart] = useState(isoDate(defaultStartDate(new Date())));
  const [end, setEnd] = useState(isoDate(new Date()));
  const [metrics, setMetrics] = useState<Metrics>({ input: true, cost: true, rate: true });
  const [busy, setBusy] = useState(false);
  const [statusText, setStatusText] = useState("");
  const [error, setError] = useState("");
  const [view, setView] = useState<FCostView | null>(null);
  const [expandedCols, setExpandedCols] = useState<Expanded>({ Day: false, Week: false, Month: false });
  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set());
  const [exporting, setExporting] = useState(false);

  useEffect(() => {
    bmesReportApi.fcost.setup().then(setSetup).catch((e: Error) => setError("Error: " + e.message));
  }, []);

  const columns = useMemo(() => (view ? visibleColumnIndexes(view, expandedCols) : []), [view, expandedCols]);
  const hierarchy = useMemo(() => (view ? buildHierarchy(view) : []), [view]);
  const byId = useMemo(() => new Map(hierarchy.map((r) => [r.rowId, r])), [hierarchy]);
  const hasKids = useMemo(() => new Set(hierarchy.filter((r) => r.parentId !== null).map((r) => r.parentId!)), [hierarchy]);
  const totals = useMemo(() => (view ? computeTotals(view, columns) : { input: [], cost: [] }), [view, columns]);

  async function generate() {
    if (busy) return;
    setBusy(true);
    setError("");
    setStatusText("");
    try {
      const result = await bmesReportApi.fcost.generate(start, end, setStatusText);
      setView(result);
      setExpandedCols({ Day: false, Week: false, Month: false });
      setSetup((s) => (s ? { ...s, rawStatus: result.rawStatus ?? s.rawStatus } : s));
      setStatusText(`Report ready: ${result.startDate} - ${result.endDate}.`);
    } catch (e) {
      setError("Error: " + (e as Error).message);
    } finally {
      setBusy(false);
    }
  }

  async function exportCsv() {
    if (exporting) return;
    setExporting(true);
    setError("");
    try {
      if (!view) {
        setError("No CSV rows to export.");
        return;
      }
      const rows = buildCsvRows(view, metrics);
      if (rows.length === 0) {
        setError("CSV export produced an empty file.");
        return;
      }
      downloadBlob(new Blob([csvText(rows)], { type: "text/csv;charset=utf-8" }), `BMES_FCost_${stamp(new Date(), true)}.csv`);
      bmesReportApi.fcost.activity("Export FCost CSV").catch(() => undefined);
    } catch (e) {
      setError("Export failed: " + (e as Error).message);
    } finally {
      setExporting(false);
    }
  }

  const toggleRow = (id: string) =>
    setExpandedRows((cur) => {
      const next = new Set(cur);
      if (!next.delete(id)) next.add(id);
      return next;
    });

  const rowsShown = metricCount(metrics);
  const cellClass = (ci: number) => (view ? colBoundaryClass(view, ci) : "");
  const credentials = setup?.credentialsConfigured ?? false;
  const foldable = view ? KINDS.filter((k) => periodColumnCount(view, k) > COLLAPSE_LIMIT[k]) : [];

  const hierarchyLabel = (row: HierarchyRow) => {
    const canExp = hasKids.has(row.rowId);
    return (
      <>
        {canExp ? <span className="fcost-tog">{expandedRows.has(row.rowId) ? "[-]" : "[+]"}</span> : <span className="fcost-tog fcost-tog-blank" />}
        <span className="hier-level-badge">{levelLabel(row)}</span>
        <strong className="ms-1">{row.display}</strong>
      </>
    );
  };

  return (
    <div className="br-page fcost">
      <div className="br-head">
        <span className="br-icon money">$</span>
        <h1>F-Cost</h1>
      </div>

      <section className="br-card">
        <div className="br-card-head"><span>Report</span></div>
        <div className="br-pad fcost-setup">
          <div className="muted small">
            Builds the report from <code>PortalGNLAB.db</code> (FCostRaw* tables). Missing dates in the selected range are fetched
            automatically before report generation; the last 3 dates are refreshed every run.
          </div>
          <div className="br-row small">
            <span className="muted">BMES:</span>
            {credentials ? (
              <>
                <span className="br-badge ok">Credentials OK</span>
                <span className="muted mono">{setup?.loginId}</span>
              </>
            ) : (
              <>
                <span className="br-badge warn">Not configured</span>
                <a href="/app/setting">Setting</a>
              </>
            )}
          </div>
          <div className="br-row">
            <label className="t5-field">
              <span className="muted small">Start</span>
              <input type="date" className="inp sm" value={start} onChange={(e) => e.target.value && setStart(e.target.value)} />
            </label>
            <label className="t5-field">
              <span className="muted small">End</span>
              <input type="date" className="inp sm" value={end} onChange={(e) => e.target.value && setEnd(e.target.value)} />
            </label>
          </div>
          <div className="br-row small">
            <span className="muted"><strong>Show rows:</strong></span>
            {([["input", "INPUT"], ["cost", "F-COST"], ["rate", "RATE %"]] as const).map(([key, label]) => (
              <label key={key} className="br-row tight">
                <input type="checkbox" checked={metrics[key]} onChange={(e) => setMetrics((m) => ({ ...m, [key]: e.target.checked }))} />
                <span>{label}</span>
              </label>
            ))}
          </div>
          <div className="small">
            <div><span className="muted">RAW DB:</span> <span className="mono break">{setup?.rawDbPath}</span></div>
            {setup && <RawStatusLine status={setup.rawStatus} />}
          </div>
          <button type="button" className="btn sm primary br-wide" onClick={generate} disabled={busy || !credentials}>
            {busy ? <><span className="br-spin" /> Loading...</> : "Generate Report"}
          </button>
          {setup && !credentials && <div className="muted small">BMES credentials are required for recent-date refresh.</div>}
          {statusText && <div className="muted small">{statusText}</div>}
        </div>
      </section>

      {view && (
        <div className="fcost-actions">
          <button type="button" className="btn sm" onClick={exportCsv} disabled={exporting}>
            {exporting ? <><span className="br-spin" /> Exporting...</> : "Export CSV"}
          </button>
        </div>
      )}

      {view && foldable.length > 0 && (
        <div className="br-row small fcost-folds">
          <span className="muted"><strong>Columns:</strong></span>
          {foldable.map((k) => (
            <button key={k} type="button" className="btn sm" onClick={() => setExpandedCols((c) => ({ ...c, [k]: !c[k] }))}>
              {k} {visiblePeriodColumnCount(view, k, expandedCols)}/{periodColumnCount(view, k)} {expandedCols[k] ? "Fold" : "Expand"}
            </button>
          ))}
        </div>
      )}

      {view && (view.subGroups.length > 0 || view.total) && (
        <section className="br-card">
          <div className="br-card-head">
            <span>Model / Sub Group Summary - Input / F-Cost / Defect Rate %</span>
            <span className="muted small">{view.subGroups.length} sub group bucket(s), {view.unmapped.length} unmapped sorted by previous-week defect rate</span>
          </div>
          <div className="scroll-x">
            <table className="fcost-table">
              <thead>
                <tr>
                  <th className="label-th" style={{ width: 260 }}>Model / Building / Line</th>
                  <th className="label-th" style={{ width: 100 }}>Type</th>
                  {columns.map((ci) => <th key={ci} className={cellClass(ci)}>{view.headers[ci]}</th>)}
                </tr>
              </thead>
              <tbody>
                {rowsShown > 0 && (
                  <>
                    <MetricRows metrics={metrics} rowClass="fcost-row-total" label={<strong>Total</strong>} columns={columns} cellClass={cellClass}
                      input={(_, i) => totals.input[i]} cost={(_, i) => totals.cost[i]} />
                    {hierarchy.filter((r) => isRowVisible(r, byId, expandedRows)).map((row) => {
                      const canExp = hasKids.has(row.rowId);
                      return (
                        <Fragment key={row.rowId}>
                          <MetricRows metrics={metrics} rowClass={row.depth === 0 ? "fcost-row-mid" : row.depth === 1 ? "fcost-row-sub1" : "fcost-row-sub2"}
                            label={hierarchyLabel(row)} labelStyle={{ paddingLeft: 8 + row.depth * 16, cursor: canExp ? "pointer" : "default" }}
                            onLabelClick={canExp ? () => toggleRow(row.rowId) : undefined}
                            columns={columns} cellClass={cellClass}
                            input={(ci) => valueAt(row.inputByCol, ci)} cost={(ci) => valueAt(row.costByCol, ci)} />
                        </Fragment>
                      );
                    })}
                  </>
                )}
              </tbody>
            </table>
          </div>
          {view.unmapped.length > 0 && (
            <div className="br-alert warn br-margin">
              <strong>Unmapped ({view.unmapped.length}):</strong> Materials whose <code>MatnrTx_ModNoTx</code> (or fallback{" "}
              <code>MatnrTx_VeridTx</code>) didn't match any LineShift in the saved ModelGroups:{" "}
              {view.unmapped.slice(0, 10).map((m) => m.displayName).join(", ")}
              {view.unmapped.length > 10 ? ` (+${view.unmapped.length - 10} more)` : ""}
              <a className="btn sm br-ml" href={unmappedModelGroupUrl(view.unmapped)} target="_blank" rel="noopener">Open Model Group for these</a>
            </div>
          )}
        </section>
      )}

      {view && (
        <section className="br-card">
          <div className="br-card-head">
            <span>F-Cost Query Date {view.queryDate}</span>
            <span className="muted small">{view.materials.length} materials / {view.totalRows} rows</span>
          </div>
          <div className="scroll-x">
            <table className="fcost-table">
              <thead>
                <tr>
                  <th className="label-th" style={{ width: 280 }}>Model</th>
                  <th className="label-th" style={{ width: 100 }}>Type</th>
                  {columns.map((ci) => <th key={ci} className={cellClass(ci)}>{view.headers[ci]}</th>)}
                </tr>
              </thead>
              <tbody>
                {rowsShown > 0 && view.subGroups.length > 0 && (
                  <MetricRows metrics={metrics} rowClass="fcost-row-total" label={<strong>{view.total?.displayName || "Total"}</strong>}
                    columns={columns} cellClass={cellClass} input={(_, i) => totals.input[i]} cost={(_, i) => totals.cost[i]} />
                )}
                {rowsShown > 0 && view.materials.map((m, idx) => (
                  <Fragment key={idx}>
                    <MetricRows metrics={metrics} rowClass="" columns={columns} cellClass={cellClass}
                      label={<><strong>{m.displayName}</strong>{m.productGroup && <div className="muted tiny">{m.productGroup}</div>}</>}
                      input={(ci) => valueAt(m.input, ci)} cost={(ci) => valueAt(m.cost, ci)} />
                  </Fragment>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {error && <div className="br-alert err">{error}</div>}
    </div>
  );
}
