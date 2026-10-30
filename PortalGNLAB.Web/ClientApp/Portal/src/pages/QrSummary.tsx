import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { openWindow, qrApi, saveBlob, type QrCell, type QrSummary as QrSummaryResult } from "../qrApi";
import { DATE_SETTLE_MS, useQrPrefs, useSettleTimer, useVirtualRows } from "./qr/hooks";
import {
  MAX_ROW_CHOICES,
  addDays,
  buildModelOptions,
  buildSummaryLines,
  clockStamp,
  compactDate,
  filterSummaryRows,
  formatPpm,
  formatStamp,
  itemFail,
  itemText,
  localToday,
  maxRowOptions,
  num0,
  scatterUrl,
  stationCards,
  stationDetailUrl,
  validModel,
  verdictClass,
  verdictLabel,
  type RowFilter,
  type SummaryLine,
} from "./qr/logic";
import "../qr.css";

const ROW_HEIGHT = 32;
const SCATTER_WINDOW = "width=1360,height=900,resizable=yes,scrollbars=yes";
const DETAIL_WINDOW = "width=1500,height=920,resizable=yes,scrollbars=yes";

function Spinner() {
  return <span className="qrs-spinner" aria-hidden="true" />;
}

function StationVerdict({ productId, cell, onOpen }: { productId: string; cell: QrCell | null; onOpen: (productId: string, cell: QrCell) => void }) {
  const title = cell ? `TestResult = ${cell.result}` : "No test record";
  return (
    <td className={`qrs-verdict ${verdictClass(cell)}`} title={title}>
      {cell ? (
        <button
          className="qrs-cell-link"
          type="button"
          title={`Open ${cell.station} data for ${productId}${cell.time ? " · " + cell.time.slice(0, 10) : ""}`}
          onClick={() => onOpen(productId, cell)}
        >
          {verdictLabel(cell)}
        </button>
      ) : (
        verdictLabel(cell)
      )}
    </td>
  );
}

function Items({ cell, count }: { cell: QrCell | null; count: number }) {
  return (
    <>
      {Array.from({ length: count }, (_, i) => (
        <td key={i} className={`qrs-value ${itemFail(cell, i) ? "qrs-fail" : ""}`}>
          {itemText(cell, i)}
        </td>
      ))}
    </>
  );
}

export default function QrSummary() {
  const { prefs, error: prefsError } = useQrPrefs();
  const [result, setResult] = useState<QrSummaryResult | null>(null);
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [maxRows, setMaxRows] = useState(2000);
  const [modelCode, setModelCode] = useState("");
  const [filter, setFilter] = useState<RowFilter>("all");
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const settle = useSettleTimer();
  const loading = useRef(false);

  const load = useCallback(async (f: string, t: string, max: number) => {
    if (loading.current) return;
    loading.current = true;
    setBusy(true);
    setError("");
    try {
      const next = await qrApi.summary(f, t, max);
      setResult(next);
      // The service resolves an empty range to the newest week; mirror what it chose.
      setFrom(next.from);
      setTo(next.to);
      setMaxRows(next.maxRows);
      setExpanded(new Set());
    } catch (e) {
      setError((e as Error).message);
    } finally {
      loading.current = false;
      setBusy(false);
    }
  }, []);

  // First load, on the cap last picked here or on the scatter.
  useEffect(() => {
    if (!prefs?.configured) return;
    const max = prefs.maxQr > 0 ? prefs.maxQr : 2000;
    setMaxRows(max);
    void load("", "", max);
  }, [prefs, load]);

  const productIds = useMemo(() => result?.rows.map((r) => r.productId) ?? [], [result]);
  const models = useMemo(() => buildModelOptions(prefs?.modelCodes ?? [], productIds), [prefs, productIds]);
  const activeModel = validModel(modelCode, models.options);
  const rows = useMemo(
    () => (result ? filterSummaryRows(result.rows, activeModel, models.known, filter) : []),
    [result, activeModel, models.known, filter],
  );
  const lines = useMemo(() => buildSummaryLines(rows, expanded), [rows, expanded]);
  const cards = useMemo(() => (result ? stationCards(result.rows) : []), [result]);
  const completeCount = result?.rows.filter((r) => r.complete).length ?? 0;
  const ngRowCount = result?.rows.filter((r) => r.hasAnyFail).length ?? 0;
  const virtual = useVirtualRows(lines.length, ROW_HEIGHT);

  const queueReload = (f: string, t: string) => settle.schedule(() => void load(f, t, maxRows), DATE_SETTLE_MS);

  const onFrom = (value: string) => {
    if (!value) return;
    setFrom(value);
    queueReload(value, to);
  };
  const onTo = (value: string) => {
    if (!value) return;
    setTo(value);
    queueReload(from, value);
  };

  /** Sets the window to the newest day with data, going back `daysBack` days. */
  const quickRange = (daysBack: number) => {
    settle.cancel();
    const anchor = result?.maxDate ?? localToday();
    setTo(anchor);
    setFrom(addDays(anchor, -daysBack));
    void load(addDays(anchor, -daysBack), anchor, maxRows);
  };

  const onMax = (value: number) => {
    if (!(value > 0)) return;
    setMaxRows(value);
    void qrApi.savePrefs({ maxQr: value }).catch(() => undefined);
    void load(from, to, value);
  };

  const toggleGroup = (productId: string) =>
    setExpanded((prev) => {
      const next = new Set(prev);
      if (!next.delete(productId)) next.add(productId);
      return next;
    });

  const openScatter = () => {
    if (!openWindow(scatterUrl(from, to, maxRows), SCATTER_WINDOW))
      setError("The browser blocked the GAUSS scatter window. Allow pop-ups for this site and try again.");
  };

  /** Opens the clicked station's data browser on this QR, one reused window per station. */
  const openDetail = (productId: string, cell: QrCell) => {
    const url = stationDetailUrl(productId, cell);
    if (url === "") return;
    if (!openWindow(url, DETAIL_WINDOW, `jino-qr-detail-${cell.station}`))
      setError(`The browser blocked the ${cell.station} data window. Allow pop-ups for this site and try again.`);
  };

  const exportExcel = async () => {
    if (exporting || rows.length === 0) return;
    setExporting(true);
    setError("");
    setNotice("");
    try {
      const { blob, truncated } = await qrApi.exportSummary(from, to);
      saveBlob(blob, `QR_SUMMARY_${compactDate(from)}-${compactDate(to)}_${clockStamp()}.xlsx`);
      if (truncated)
        setNotice("The export stops at 50,000 QR, which this period reached. Narrow the date range to get the whole period in one file.");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setExporting(false);
    }
  };

  const configured = prefs?.configured ?? true;
  const connection = prefs
    ? `${prefs.connection.server},${prefs.connection.port} / ${prefs.connection.database} / ALTD·BKTD·GSTD`
    : "";
  const body = (line: SummaryLine) => {
    const row = line.row;
    const cls = [line.isChild ? "qrs-row-child" : "", row.hasAnyFail ? "qrs-row-ng" : ""].filter(Boolean).join(" ");
    return (
      <tr key={`${row.productId}:${line.isChild ? row.round : "top"}`} className={cls}>
        <td className="qrs-group-cell">
          {line.isChild ? (
            <span className="qrs-child-mark" aria-label="Earlier test round for the same QR">└</span>
          ) : row.older.length > 0 ? (
            <button
              className="qrs-group-toggle"
              type="button"
              aria-expanded={line.expanded}
              title={
                line.expanded
                  ? `Collapse the ${num0(row.older.length)} earlier test rounds of this QR`
                  : `${num0(row.older.length)} earlier test rounds hidden — click to expand, newest first`
              }
              onClick={() => toggleGroup(row.productId)}
            >
              <span aria-hidden="true">{line.expanded ? "−" : "+"}</span>
              <b>{num0(row.older.length)}</b>
            </button>
          ) : null}
        </td>
        <td className="qrs-qr">{line.isChild ? "" : row.productId}</td>
        <td className="qrs-time">{formatStamp(row.latestTime)}</td>
        <td className={`qrs-verdict ${row.finalPass ? "qrs-pass" : "qrs-fail"}`} title={row.finalNote}>
          {row.finalPass ? "PASS" : "NG"}
        </td>
        <StationVerdict productId={row.productId} cell={row.airleak} onOpen={openDetail} />
        <Items cell={row.airleak} count={1} />
        <td className="qrs-time">{formatStamp(row.airleak?.time)}</td>
        <StationVerdict productId={row.productId} cell={row.bako} onOpen={openDetail} />
        <Items cell={row.bako} count={5} />
        <td className="qrs-time">{formatStamp(row.bako?.time)}</td>
        <StationVerdict productId={row.productId} cell={row.gauss} onOpen={openDetail} />
        <Items cell={row.gauss} count={2} />
        <td className="qrs-time">{formatStamp(row.gauss?.time)}</td>
      </tr>
    );
  };

  return (
    <div className="qrs-page">
      <header className="qrs-titlebar">
        <div>
          <div className="qrs-kicker">AIRLEAK · BAKO · GAUSS COMBINED</div>
          <h1>QR SUMMARY</h1>
          <p>One row per Product ID across the three stations. Each value is that station’s most recent test within the selected period.</p>
        </div>
        <div className="qrs-actions">
          <span className="qrs-connection">{connection}</span>
          <button className="qrs-btn qrs-primary" type="button" onClick={() => { settle.cancel(); void load(from, to, maxRows); }} disabled={busy || !configured}>
            {busy ? (<><Spinner /><span>Loading</span></>) : <span>Refresh</span>}
          </button>
          <button
            className="qrs-btn"
            type="button"
            title="Two sheets: the newest round per QR, and every round of the selected period"
            onClick={() => void exportExcel()}
            disabled={busy || exporting || rows.length === 0}
          >
            {exporting ? (<><Spinner /><span>Exporting</span></>) : <span>XLSX</span>}
          </button>
        </div>
      </header>

      {(error || prefsError) && <div className="qrs-alert qrs-danger">{error || prefsError}</div>}
      {notice && <div className="qrs-alert qrs-warning">{notice}</div>}
      {prefs && !prefs.configured ? (
        <div className="qrs-alert qrs-warning">QR DB password is not configured on the server.</div>
      ) : (
        result?.warning && <div className="qrs-alert qrs-warning">{result.warning}</div>
      )}

      <div className="qrs-toolbar">
        <label className="qrs-field">
          <span>From</span>
          <input type="date" value={from} onChange={(e) => onFrom(e.target.value)} />
        </label>
        <label className="qrs-field">
          <span>To</span>
          <input type="date" value={to} onChange={(e) => onTo(e.target.value)} />
        </label>
        <div className="qrs-quick">
          <button className="qrs-chip" type="button" disabled={busy} onClick={() => quickRange(0)}>Latest day</button>
          <button className="qrs-chip" type="button" disabled={busy} onClick={() => quickRange(6)}>7 days</button>
          <button className="qrs-chip" type="button" disabled={busy} onClick={() => quickRange(29)}>30 days</button>
        </div>
        <label className="qrs-field">
          <span>Max QR</span>
          <select disabled={busy} value={maxRows} onChange={(e) => onMax(Number(e.target.value))}>
            {maxRowOptions(MAX_ROW_CHOICES, maxRows).map((option) => (
              <option key={option} value={option}>{num0(option)}</option>
            ))}
          </select>
        </label>
        <label className="qrs-field qrs-grow">
          <span>Model</span>
          <select disabled={busy} value={activeModel} onChange={(e) => setModelCode(e.target.value)}>
            <option value="">All models</option>
            {models.options.map((option) => (
              <option key={option.code} value={option.code}>{option.label}</option>
            ))}
          </select>
        </label>
        <div className="qrs-quick">
          <button className={`qrs-chip ${filter === "all" ? "qrs-on" : ""}`} type="button" onClick={() => setFilter("all")}>All</button>
          <button className={`qrs-chip ${filter === "ng" ? "qrs-on" : ""}`} type="button" onClick={() => setFilter("ng")}>Has NG</button>
          <button className={`qrs-chip ${filter === "complete" ? "qrs-on" : ""}`} type="button" onClick={() => setFilter("complete")}>All 3 stations</button>
        </div>
      </div>

      {busy && !result ? (
        <div className="qrs-state"><Spinner /><span>Reading ALTD / BKTD / GSTD from TCDB.</span></div>
      ) : !result ? (
        <div className="qrs-state">No data loaded.</div>
      ) : (
        <>
          <div className="qrs-cards">
            <article className="qrs-card">
              <span>QR count</span>
              <strong>{num0(result.rows.length)}</strong>
              <small>{result.hitMaxRows ? `Latest ${num0(result.maxRows)} of ${num0(result.totalProducts)} total` : "Whole selected period"}</small>
            </article>
            {cards.map((card) =>
              card.station === "GAUSS" ? (
                <button key={card.station} className="qrs-card qrs-card-link" type="button" title="Open the GAUSS dispersion scatter in a new window" onClick={openScatter}>
                  <span>{card.station} <em className="qrs-card-cue">SCATTER ↗</em></span>
                  <strong>{num0(card.tested)}</strong>
                  <small>NG {num0(card.ng)} · {formatPpm(card.ngRatePpm)}</small>
                </button>
              ) : (
                <article key={card.station} className="qrs-card">
                  <span>{card.station}</span>
                  <strong>{num0(card.tested)}</strong>
                  <small>NG {num0(card.ng)} · {formatPpm(card.ngRatePpm)}</small>
                </article>
              ),
            )}
            <article className="qrs-card">
              <span>All 3 stations tested</span>
              <strong>{num0(completeCount)}</strong>
              <small>QR with data from all three stations</small>
            </article>
            <article className="qrs-card">
              <span>QR with NG</span>
              <strong>{num0(ngRowCount)}</strong>
              <small>QR failing at least one station</small>
            </article>
          </div>

          <div className="qrs-table-wrap" ref={virtual.ref} onScroll={virtual.onScroll}>
            {rows.length === 0 ? (
              <div className="qrs-empty">No QR matches the current filter.</div>
            ) : (
              <table className="qrs-table">
                <thead>
                  <tr className="qrs-group-row">
                    <th rowSpan={2} className="qrs-group-head">QR GROUP</th>
                    <th rowSpan={2} className="qrs-qr-head">Product ID</th>
                    <th rowSpan={2} className="qrs-time-head">Last test</th>
                    <th rowSpan={2} className="qrs-final-head">Final Result</th>
                    <th colSpan={3} className="qrs-grp qrs-grp-air">AIRLEAK · dbo.ALTD</th>
                    <th colSpan={7} className="qrs-grp qrs-grp-bako">BAKO · dbo.BKTD</th>
                    <th colSpan={4} className="qrs-grp qrs-grp-gauss">
                      <button className="qrs-grp-link" type="button" title="Open the GAUSS dispersion scatter in a new window" onClick={openScatter}>
                        GAUSS · dbo.GSTD ↗
                      </button>
                    </th>
                  </tr>
                  <tr>
                    <th>Result</th><th>Value</th><th>Test time</th>
                    <th>Result</th><th>SPL1</th><th>SPL2</th><th>THD</th><th>ImpLimit</th><th>Imp2kHz</th><th>Test time</th>
                    <th>Result</th><th>Value</th><th>Value 2</th><th>Test time</th>
                  </tr>
                </thead>
                <tbody>
                  <tr aria-hidden="true"><td colSpan={18} style={{ height: virtual.before, padding: 0, border: 0 }} /></tr>
                  {lines.slice(virtual.first, virtual.last).map(body)}
                  <tr aria-hidden="true"><td colSpan={18} style={{ height: virtual.after, padding: 0, border: 0 }} /></tr>
                </tbody>
              </table>
            )}
          </div>

          <footer className="qrs-foot">
            <span>{result.from} ~ {result.to}</span>
            <span>Shown {num0(rows.length)} / fetched {num0(result.rows.length)} QR</span>
            <span>Only 17-character Product IDs are counted</span>
            {result.fetchedAt && <span>as of {result.fetchedAt}</span>}
          </footer>
        </>
      )}
    </div>
  );
}
