import { useMemo, useState } from "react";
import { bmesApi, type WorkerFetchResult } from "../bmesApi";
import { downloadBlob } from "../toolsApi";
import "../bmes.css";
import {
  METRICS, buildDepartments, checkInStyle, dateColumns, isLate, normalizeRange, statusClass, summarize, weekday, type Metric,
} from "./bmes/workerLogic";
import type { WorkerRecord } from "../bmesApi";

const today = () => {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
};

const exportName = (start: string, end: string) =>
  start === end ? `WorkerStatus_${start.replaceAll("-", "")}.xlsx` : `WorkerStatus_${start.replaceAll("-", "")}_${end.replaceAll("-", "")}.xlsx`;

const dash = <span className="ws-empty-cell">-</span>;

function MetricCell({ metric, rec }: { metric: Metric; rec: WorkerRecord | undefined }) {
  if (!rec) return dash;
  switch (metric) {
    case "Day":
      return <span>{weekday(rec.date.slice(0, 10))}</span>;
    case "Status":
      return rec.workStatus ? <span className={"ws-badge " + statusClass(rec.workStatus)}>{rec.workStatus}</span> : dash;
    case "Date Type":
      return rec.dayType ? <span>{rec.dayType}</span> : dash;
    case "Check In":
      return (
        <>
          {rec.checkIn ? <span className="ws-checkin" style={checkInStyle(rec)}>{isLate(rec) ? "*" : ""}{rec.checkIn}</span> : dash}
          {rec.schedStart && <span className="ws-sched">({rec.schedStart})</span>}
        </>
      );
    case "Check Out":
      return (
        <>
          {rec.checkOut || dash}
          {rec.schedEnd && <span className="ws-sched">({rec.schedEnd})</span>}
        </>
      );
  }
}

export default function BmesWorkerStatus() {
  const [start, setStart] = useState(today());
  const [end, setEnd] = useState(today());
  const [busy, setBusy] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [progress, setProgress] = useState("");
  const [result, setResult] = useState<WorkerFetchResult | null>(null);
  const [fetchedRange, setFetchedRange] = useState({ start: "", end: "" });
  const [collapsed, setCollapsed] = useState<Set<string>>(new Set());

  const ok = result?.isSuccess === true;
  const records = result?.records ?? [];
  const hasRows = ok && records.length > 0;
  const summary = useMemo(() => summarize(records, result?.totalCount ?? 0), [records, result]);
  const depts = useMemo(() => buildDepartments(records), [records]);
  const dates = useMemo(() => dateColumns(records), [records]);

  const changeStart = (value: string) => {
    if (!value) return;
    const range = normalizeRange(value, end, "start");
    setStart(range.start);
    setEnd(range.end);
  };
  const changeEnd = (value: string) => {
    if (!value) return;
    const range = normalizeRange(start, value, "end");
    setStart(range.start);
    setEnd(range.end);
  };

  async function fetchRange() {
    setBusy(true);
    setProgress("");
    setResult(null);
    setCollapsed(new Set());
    try {
      const data = await bmesApi.worker.fetch(start, end, setProgress);
      setResult(data);
      setFetchedRange({ start, end });
    } catch (e) {
      setResult({ isSuccess: false, errorMessage: (e as Error).message, totalCount: 0, fetchedAt: "", records: [] });
    } finally {
      setBusy(false);
    }
  }

  async function exportExcel() {
    if (exporting || !hasRows) return;
    setExporting(true);
    setProgress("");
    try {
      const blob = await bmesApi.worker.exportXlsx(fetchedRange.start, fetchedRange.end, records);
      downloadBlob(blob, exportName(fetchedRange.start, fetchedRange.end));
    } catch (e) {
      setProgress(`Export failed: ${(e as Error).message}`);
    } finally {
      setExporting(false);
    }
  }

  const toggle = (key: string) => {
    const next = new Set(collapsed);
    if (!next.delete(key)) next.add(key);
    setCollapsed(next);
  };

  const fetchedAt = result?.fetchedAt ? new Date(result.fetchedAt).toLocaleTimeString("en-GB") : "";
  let personNumber = 0;

  return (
    <div className="bm-page">
      <div className="bm-head"><h1>Worker Status</h1></div>

      <div className="ws-card">
        <div className="ws-query-row">
          <label>From</label>
          <input type="date" className="bm-input" style={{ width: 160 }} value={start} disabled={busy} onChange={(e) => changeStart(e.target.value)} />
          <label>To</label>
          <input type="date" className="bm-input" style={{ width: 160 }} value={end} disabled={busy} onChange={(e) => changeEnd(e.target.value)} />
          <button className="btn sm primary" onClick={() => void fetchRange()} disabled={busy}>{busy ? "Fetching…" : "Fetch"}</button>
          {hasRows && (
            <button className="btn sm" onClick={() => void exportExcel()} disabled={busy || exporting} title="Download as Excel (.xlsx)">
              {exporting ? "Exporting…" : "Export Excel"}
            </button>
          )}
          {ok && fetchedAt && <span className="ws-fetched">fetched {fetchedAt}</span>}
        </div>
        {progress && <div className="ws-progress">{progress}</div>}
        {result?.errorMessage && <div className="ws-error" role="alert">⚠ {result.errorMessage}</div>}
      </div>

      {hasRows && (
        <div className="ws-summary">
          <div className="ws-pill">Rows <b>{summary.rows}</b></div>
          <div className="ws-pill">People <b>{summary.people}</b></div>
          <div className="ws-pill">Days <b>{summary.days}</b></div>
          <div className="ws-pill in">Checked In <b>{summary.checkedIn}</b></div>
          <div className="ws-pill in">Checked Out <b>{summary.checkedOut}</b></div>
          {summary.late > 0 && <div className="ws-pill late">Late <b>{summary.late}</b></div>}
        </div>
      )}

      <div className="ws-card">
        <div className="ws-card-title">
          <span>Attendance</span>
          {ok && <span className="count">{result!.totalCount} records</span>}
        </div>

        {!hasRows ? (
          <div className="ws-empty">{result === null ? "Pick a date and click Fetch." : !ok ? "Fetch failed." : "No records found."}</div>
        ) : (
          <div className="ws-table-wrap">
            <table className="ws-table">
              <thead>
                <tr>
                  <th>#</th><th>Dept Code</th><th>Dept No</th><th>Emp No</th><th>Name</th><th style={{ width: 90 }}>Metric</th>
                  {dates.length > 0 && <th colSpan={dates.length} style={{ background: "#f8fafc" }}></th>}
                </tr>
              </thead>
              <tbody>
                {depts.map((dept) => {
                  const closed = collapsed.has(dept.key);
                  return [
                    <tr key={dept.key} className="ws-dept-row">
                      <td colSpan={6 + dates.length}>
                        <button type="button" className="ws-dept-toggle" title="Toggle department" onClick={() => toggle(dept.key)}>{closed ? "▸" : "▾"}</button>
                        <span>{dept.code}</span>
                        <span className="ws-dept-meta">{dept.no}</span>
                        {dept.name.trim() && <span className="ws-dept-meta">{dept.name}</span>}
                        <span className="ws-dept-meta">{dept.people} people / {dept.rows} rows</span>
                      </td>
                    </tr>,
                    ...(closed ? [] : [
                      <tr key={dept.key + "-dates"} className="ws-date-row">
                        <td colSpan={5}></td>
                        <td className="ws-metric-label">Date</td>
                        {dates.map((d) => <th key={d} className="ws-date-th">{d}</th>)}
                      </tr>,
                      ...dept.persons.flatMap((person) => {
                        personNumber++;
                        const first = person.records[0];
                        return METRICS.map((metric, mi) => (
                          <tr key={`${dept.key}-${person.key}-${metric}`} style={mi === 0 && personNumber > 1 ? { borderTop: "2px solid #e2e8f0" } : undefined}>
                            {mi === 0 && (
                              <>
                                <td rowSpan={METRICS.length} className="ws-row-number">{personNumber}</td>
                                <td rowSpan={METRICS.length} className="ws-code" title={first.departmentName}>{first.departmentCode}</td>
                                <td rowSpan={METRICS.length} className="ws-code">{first.departmentNo}</td>
                                <td rowSpan={METRICS.length} className="ws-code">{first.empNo}</td>
                                <td rowSpan={METRICS.length} style={{ fontWeight: 600, verticalAlign: "top" }}>{first.name}</td>
                              </>
                            )}
                            <td className="ws-metric-label">{metric}</td>
                            {dates.map((d) => (
                              <td key={d} className={metric === "Check In" || metric === "Check Out" ? "ws-time" : undefined}>
                                <MetricCell metric={metric} rec={person.byDate.get(d)} />
                              </td>
                            ))}
                          </tr>
                        ));
                      }),
                    ]),
                  ];
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
