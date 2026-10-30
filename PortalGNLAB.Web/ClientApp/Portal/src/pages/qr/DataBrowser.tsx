import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { qrApi, saveBlob, type QrColumn, type QrData } from "../../qrApi";
import { useQrPrefs, useSettleTimer, useVirtualRows } from "./hooks";
import {
  DATA_MAX_ROW_CHOICES,
  TEST_DATE_INDEX,
  buildDataLines,
  buildModelOptions,
  buildNgBreakdown,
  buildNgItemColumns,
  buildRowGroups,
  cellValue,
  clockStamp,
  displayCell,
  exportDateText,
  filterGroups,
  filteredSummary,
  findColumnIndex,
  formatDay,
  formatPpm,
  isFiltered,
  makeJudge,
  maxRowOptions,
  num0,
  parseDay,
  rowElementId,
  selectedDatesLabel,
  validModel,
  type DataFilters,
  type DataLine,
} from "./logic";
import "../../qr.css";

const ROW_HEIGHT = 32;
const FILTER_SETTLE_MS = 220;

export interface DataBrowserProps {
  /** BKTD / ALTD / GSTD */
  tableKey: string;
  kicker: string;
  description: string;
}

function Spinner() {
  return <span className="qr-spinner" aria-hidden="true" />;
}

function cellClass(column: QrColumn, value: string, judge: ReturnType<typeof makeJudge>): string {
  const numeric = column.numeric ? "qr-num" : "";
  if (!column.name.toLowerCase().endsWith("result")) return numeric;
  const result = judge.isFail(value) ? "qr-result-fail" : judge.isPass(value) ? "qr-result-pass" : "";
  return [numeric, result].filter(Boolean).join(" ");
}

/** Browser of one TCDB test table (BKTD / ALTD / GSTD): date multi-select, QR grouping, column filters, NG breakdown. */
export default function DataBrowser({ tableKey, kicker, description }: DataBrowserProps) {
  const { prefs, error: prefsError } = useQrPrefs();
  // Deep link from QR SUMMARY: one QR on one date.
  const focus = useMemo(() => {
    const params = new URLSearchParams(window.location.search);
    return { date: parseDay(params.get("date")), productId: (params.get("qr") ?? "").trim() };
  }, []);

  const [data, setData] = useState<QrData | null>(null);
  const [selected, setSelected] = useState<string[]>([]);
  const [maxRows, setMaxRows] = useState(0);
  const [busy, setBusy] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [filters, setFilters] = useState<DataFilters>({ search: focus.productId, modelCode: "", columns: {} });
  const [typed, setTyped] = useState<Record<number, string>>({});
  const [epoch, setEpoch] = useState(0);
  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [expandedNg, setExpandedNg] = useState<Set<string>>(new Set());
  const [focusRowId, setFocusRowId] = useState("");
  // What the server resolved Auto to, shown next to the select.
  const [autoCap, setAutoCap] = useState(0);
  const settle = useSettleTimer();
  const loading = useRef(false);
  const scrollPending = useRef(false);
  const focusScrolled = useRef(false);

  const load = useCallback(
    async (dates: string[], cap: number) => {
      if (loading.current) return;
      loading.current = true;
      setBusy(true);
      setError("");
      try {
        const next = await qrApi.data(tableKey, dates, cap);
        setData(next);
        setSelected(next.selectedDates);
        setMaxRows(cap <= 0 ? 0 : next.maxRows);
        setAutoCap(next.maxRows);
        setExpanded(new Set());
        // A deep link names one QR on one date; unfold it and scroll it into view.
        if (focus.productId !== "") {
          const groups = buildRowGroups(next);
          const match = groups.find((g) => g.key.toLowerCase() === `product:${focus.productId}`.toLowerCase());
          if (!match) {
            setFocusRowId("");
            setError(`${focus.productId} has no row on the selected date in ${next.table.displayTable}.`);
          } else {
            if (match.rows.length > 1) setExpanded(new Set([match.key]));
            setFocusRowId(rowElementId(match.key));
            if (!focusScrolled.current) {
              focusScrolled.current = true;
              scrollPending.current = true;
            }
          }
        }
      } catch (e) {
        setError((e as Error).message);
      } finally {
        loading.current = false;
        setBusy(false);
      }
    },
    [tableKey, focus.productId],
  );

  const auto = maxRows <= 0;

  // First load: Auto unless a fixed cap was picked on one of the QR data pages.
  useEffect(() => {
    if (!prefs?.configured) return;
    const cap = prefs.maxRows > 0 ? prefs.maxRows : 0;
    setMaxRows(cap);
    void load(focus.date ? [focus.date] : [], cap);
  }, [prefs, load, focus.date]);

  const judge = useMemo(() => makeJudge(data?.table.passValues ?? [], data?.table.failValues ?? []), [data]);
  const groups = useMemo(() => (data ? buildRowGroups(data) : []), [data]);
  const testResultIndex = useMemo(() => (data ? findColumnIndex(data.columns, "TestResult") : -1), [data]);
  const productIdIndex = useMemo(() => (data ? findColumnIndex(data.columns, "ProductID") : -1), [data]);
  const ngItemColumns = useMemo(() => (data ? buildNgItemColumns(data.columns) : []), [data]);
  const models = useMemo(
    () =>
      productIdIndex < 0
        ? { options: [], known: new Set<string>() }
        : buildModelOptions(prefs?.modelCodes ?? [], groups.map((g) => cellValue(g.rows[0], productIdIndex))),
    [prefs, groups, productIdIndex],
  );
  const modelCode = validModel(filters.modelCode, models.options);
  const active: DataFilters = useMemo(() => ({ ...filters, modelCode }), [filters, modelCode]);
  const filtered = isFiltered(active);
  const visible = useMemo(() => filterGroups(groups, active, productIdIndex, models.known), [groups, active, productIdIndex, models.known]);
  const lines = useMemo(() => buildDataLines(visible, expanded), [visible, expanded]);
  const breakdown = useMemo(() => buildNgBreakdown(visible, ngItemColumns, testResultIndex, judge), [visible, ngItemColumns, testResultIndex, judge]);
  const summary = filtered ? (data?.summary ? filteredSummary(visible, testResultIndex, judge) : null) : (data?.summary ?? null);
  const virtual = useVirtualRows(lines.length, ROW_HEIGHT);

  // Scroll the deep-linked QR into view once, after the first render that holds it.
  useEffect(() => {
    if (!scrollPending.current || !focusRowId) return;
    const index = lines.findIndex((l) => !l.isChild && rowElementId(l.group.key) === focusRowId);
    const el = virtual.ref.current;
    if (index < 0 || !el) return;
    scrollPending.current = false;
    el.scrollTop = Math.max(index * ROW_HEIGHT - el.clientHeight / 2, 0);
  }, [lines, focusRowId, virtual.ref]);

  const dateColumn = data && data.sortColumnName.trim() !== "" ? data.sortColumnName : "TestTime";
  const hitMax = data?.hitMaxRows === true;

  const toggleDate = (date: string) => {
    if (busy) return;
    const next = selected.includes(date) ? selected.filter((d) => d !== date) : [...selected, date];
    void load(next, maxRows);
  };

  const onMaxRows = (value: number) => {
    if (busy || value === maxRows) return;
    void qrApi.savePrefs({ maxRows: value <= 0 ? -1 : value }).catch(() => undefined);
    setMaxRows(value);
    void load(selected, value);
  };

  const setColumnFilter = (index: number, text: string) => {
    setTyped((t) => ({ ...t, [index]: text }));
    settle.schedule(
      () => setFilters((f) => ({ ...f, columns: { ...f.columns, [index]: text } })),
      FILTER_SETTLE_MS,
    );
  };

  const clearColumnFilter = (index: number) => {
    settle.cancel();
    setTyped((t) => ({ ...t, [index]: "" }));
    setFilters((f) => ({ ...f, columns: { ...f.columns, [index]: "" } }));
  };

  const showAll = () => {
    settle.cancel();
    setTyped({});
    setEpoch((e) => e + 1);
    setFilters({ search: "", modelCode: "", columns: {} });
  };

  const toggleGroup = (key: string) =>
    setExpanded((prev) => {
      const next = new Set(prev);
      if (!next.delete(key)) next.add(key);
      return next;
    });

  const toggleNg = (label: string) =>
    setExpandedNg((prev) => {
      const next = new Set(prev);
      if (!next.delete(label)) next.add(label);
      return next;
    });

  const exportExcel = async () => {
    if (exporting || !data || data.rows.length === 0) return;
    setExporting(true);
    setError("");
    setNotice("");
    try {
      const { blob, truncated } = await qrApi.exportData(tableKey, selected);
      saveBlob(blob, `QR_${data.table.station}_DATA_${data.table.key}_${exportDateText(selected)}_${clockStamp()}.xlsx`);
      if (truncated)
        setNotice("The export stops at 100,000 rows, which the selected dates reached. Select fewer dates to get the whole period in one file.");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setExporting(false);
    }
  };

  const columnLabel = (index: number) =>
    index === TEST_DATE_INDEX ? "Test date" : data && index >= 0 && index < data.columns.length ? data.columns[index].name : `#${index}`;
  const activeColumnFilters = Object.entries(active.columns)
    .filter(([, text]) => text.trim() !== "")
    .map(([key, text]) => ({ index: Number(key), text: text.trim() }))
    .sort((a, b) => a.index - b.index);

  const configured = prefs?.configured ?? true;
  const title = data?.table.menuLabel ?? tableKey;
  const connection = prefs ? `${prefs.connection.server},${prefs.connection.port} / ${prefs.connection.database} / ${prefs.connection.userId}` : "";
  const colCount = (data?.layout.length ?? 0) + 1;

  const renderLine = (line: DataLine, i: number) => {
    const key = line.group.key;
    const cls = line.isChild ? "qr-child-row" : rowElementId(key) === focusRowId ? "qr-primary-row qr-focus-row" : "qr-primary-row";
    return (
      <tr key={`${key}:${virtual.first + i}`} className={cls} id={line.isChild ? undefined : rowElementId(key)}>
        <td className="qr-expand-cell">
          {line.isChild ? (
            <span className="qr-child-mark" aria-label="Additional test row for the same QR">└</span>
          ) : line.group.rows.length > 1 ? (
            <button
              className="qr-group-toggle"
              type="button"
              aria-expanded={line.expanded}
              title={line.expanded ? "Collapse rows for this QR" : "Expand rows for this QR"}
              onClick={() => toggleGroup(key)}
            >
              <span aria-hidden="true">{line.expanded ? "−" : "+"}</span>
              <b>{num0(line.group.rows.length)}</b>
            </button>
          ) : null}
        </td>
        {data!.layout.map((slot) => {
          if (slot === TEST_DATE_INDEX) return <td key="date" className="qr-test-date">{line.row.date ?? "-"}</td>;
          const value = cellValue(line.row, slot);
          return (
            <td key={slot} className={cellClass(data!.columns[slot], value, judge)}>
              {displayCell(value)}
            </td>
          );
        })}
      </tr>
    );
  };

  const filterBox = (index: number, label: string) => (
    <input
      key={`qr-cf-${index}-${epoch}`}
      className={`qr-col-filter ${(active.columns[index] ?? "").trim() !== "" ? "qr-on" : ""}`}
      type="search"
      placeholder="filter"
      aria-label={`Filter by ${label}`}
      value={typed[index] ?? ""}
      onChange={(e) => setColumnFilter(index, e.target.value)}
    />
  );

  return (
    <div className="qr-date-browser">
      <header className="qr-titlebar">
        <div>
          <div className="qr-kicker">{kicker}</div>
          <h1>{title}</h1>
          <p>{description}</p>
        </div>
        <div className="qr-actions">
          <span className="qr-connection">{connection}</span>
          <button className="qr-btn qr-primary" type="button" onClick={() => void load(selected, maxRows)} disabled={busy || !configured}>
            {busy ? (<><Spinner /><span>Loading</span></>) : <span>Refresh</span>}
          </button>
          <button
            className="qr-btn"
            type="button"
            title="Two sheets: the newest test per QR, and every row of the selected dates"
            onClick={() => void exportExcel()}
            disabled={busy || exporting || !data || data.rows.length === 0}
          >
            {exporting ? (<><Spinner /><span>Exporting</span></>) : <span>XLSX</span>}
          </button>
        </div>
      </header>

      {(error || prefsError) && <div className="qr-alert qr-danger">{error || prefsError}</div>}
      {notice && <div className="qr-alert qr-warning">{notice}</div>}
      {data?.warning && <div className="qr-alert qr-warning">{data.warning}</div>}
      {prefs && !prefs.configured && <div className="qr-alert qr-warning">QR DB password is not configured on the server.</div>}

      {busy && !data ? (
        <div className="qr-state-panel"><Spinner /><span>Loading the date list from TCDB.{tableKey}</span></div>
      ) : !data ? (
        <div className="qr-state-panel">No data loaded.</div>
      ) : (
        <div className="qr-browser-grid" aria-busy={busy}>
          <aside className="qr-date-panel">
            <div className="qr-section-head">
              <div>
                <strong>Test dates</strong>
                <span>{num0(selected.length)} selected</span>
              </div>
              <small>{dateColumn} · multi-select</small>
            </div>
            <div className="qr-date-list">
              {data.dates.length === 0 ? (
                <div className="qr-empty">No dates to show.</div>
              ) : (
                data.dates.map((d) => (
                  <label key={d.date} className={`qr-date-option${selected.includes(d.date) ? " qr-selected" : ""}${busy ? " qr-disabled" : ""}`}>
                    <input className="qr-date-check" type="checkbox" checked={selected.includes(d.date)} disabled={busy} onChange={() => toggleDate(d.date)} />
                    <span className="qr-date-label">{formatDay(d.date)}</span>
                    <span className="qr-date-count">{num0(d.rowCount)} rows</span>
                  </label>
                ))
              )}
            </div>
          </aside>

          <section className="qr-data-panel">
            <div className="qr-section-head qr-table-head">
              <div>
                <strong>{selectedDatesLabel(selected)}</strong>
                <span>Test data</span>
              </div>
              <div className="qr-row-summary">
                <span>Shown {num0(visible.length)} QR{filtered ? ` of ${num0(groups.length)}` : ""}</span>
                <span>Fetched {num0(data.rows.length)} rows</span>
                {data.totalRows != null && <span>DB total {num0(data.totalRows)} rows</span>}
                {data.excludedRows != null && (
                  <span className={data.excludedRows > 0 ? "qr-excluded" : undefined}>Excluded {num0(data.excludedRows)} rows</span>
                )}
                <span>{num0(data.columns.length)} columns</span>
                {data.validTotalRows != null && data.validTotalRows > data.rows.length && (
                  <b>Limited to the latest {num0(data.maxRows)} rows · QR groups follow the shown range</b>
                )}
              </div>
            </div>

            {summary && (
              <div className={filtered ? "qr-summary-grid qr-summary-filtered" : "qr-summary-grid"} aria-label={filtered ? "QR summary for the filtered rows" : "QR summary for the selected dates"}>
                <article className="qr-summary-card">
                  <span>Input count{filtered ? " · filtered" : ""}</span>
                  <strong>{num0(summary.inputCount)}</strong>
                  <small>{filtered ? "Current filter · duplicate ProductID excluded" : hitMax ? "Shown range · duplicate ProductID excluded" : "Duplicate ProductID excluded"}</small>
                </article>
                <article className="qr-summary-card">
                  <span>NG count{filtered ? " · filtered" : ""}</span>
                  <strong>{num0(summary.ngCount)}</strong>
                  <small>{filtered ? `Current filter · latest ${dateColumn}` : hitMax ? `Shown range · latest ${dateColumn}` : `Latest ${dateColumn} result`}</small>
                </article>
                <article className="qr-summary-card">
                  <span>NG Rate{filtered ? " · filtered" : ""}</span>
                  <strong>{formatPpm(summary.ngRatePpm)}</strong>
                  <small>NG ÷ Input × 1,000,000 (ppm)</small>
                </article>
              </div>
            )}

            {(breakdown.groups.length > 0 || breakdown.unspecified > 0) && (
              <div className="qr-ng-items" aria-label="NG items, unfold one to see its combinations">
                <div className="qr-ng-head">
                  <strong>NG by item</strong>
                  <span>
                    {num0(breakdown.rowCount)} NG QR shown · {num0(breakdown.itemTotalSum)} item NG in {num0(breakdown.comboCount)} combinations
                    {breakdown.multiCount > 0 && <b> · {num0(breakdown.multiCount)} failed 2+ items</b>}
                    {" "}· click an item to unfold what it failed with
                  </span>
                </div>
                <div className="qr-ng-list">
                  {breakdown.groups.map((group) => {
                    const open = expandedNg.has(group.label);
                    return (
                      <span key={group.label} className={open ? "qr-ng-group qr-ng-open" : "qr-ng-group"}>
                        <button
                          className="qr-ng-chip qr-ng-parent"
                          type="button"
                          aria-expanded={open}
                          title={`${group.label} failed on ${num0(group.total)} of the QR shown, alone or with other items`}
                          onClick={() => toggleNg(group.label)}
                        >
                          <span aria-hidden="true">{open ? "−" : "+"}</span>
                          <b>{group.label}</b>
                          <i>{num0(group.total)}</i>
                        </button>
                        {open &&
                          group.children.map((child) => (
                            <span
                              key={child.label}
                              className="qr-ng-chip qr-ng-child"
                              title={child.label === "alone" ? `${num0(child.count)} QR failed ${group.label} and nothing else` : `${num0(child.count)} QR failed ${group.label} together with ${child.label.slice(2)}`}
                            >
                              <b>{child.label}</b>
                              <i>{num0(child.count)}</i>
                            </span>
                          ))}
                      </span>
                    );
                  })}
                  {breakdown.unspecified > 0 && (
                    <span className="qr-ng-chip qr-ng-other" title="Failed the unit verdict without any item column flagged">
                      <b>Unspecified</b>
                      <i>{num0(breakdown.unspecified)}</i>
                    </span>
                  )}
                </div>
              </div>
            )}

            <div className="qr-search-bar">
              <label className="qr-maxrows">
                <span>Max rows</span>
                <select disabled={busy} value={maxRows} onChange={(e) => onMaxRows(Number(e.target.value))}>
                  <option value={0}>Auto · fit the dates</option>
                  {maxRowOptions(DATA_MAX_ROW_CHOICES, auto ? 0 : maxRows).map((option) => (
                    <option key={option} value={option}>{num0(option)}</option>
                  ))}
                </select>
                {auto && <small className="qr-maxrows-auto">{num0(data.maxRows || autoCap)}</small>}
              </label>

              <label className="qr-model-field">
                <span>Model</span>
                <select disabled={busy || productIdIndex < 0} value={modelCode} onChange={(e) => setFilters((f) => ({ ...f, modelCode: e.target.value }))}>
                  <option value="">All models</option>
                  {models.options.map((option) => (
                    <option key={option.code} value={option.code}>{option.label}</option>
                  ))}
                </select>
                {productIdIndex < 0 ? (
                  <small className="qr-model-note">no ProductID column</small>
                ) : (
                  models.options.length === 0 && <small className="qr-model-note">Setting → QR Model Codes is empty</small>
                )}
              </label>

              {activeColumnFilters.map(({ index, text }) => (
                <button key={index} className="qr-filter-chip" type="button" title="Clear this column filter" onClick={() => clearColumnFilter(index)}>
                  <b>{columnLabel(index)}</b>
                  <span>{text}</span>
                  <i aria-hidden="true">×</i>
                </button>
              ))}
              {filtered && <button className="qr-btn qr-search-clear" type="button" onClick={showAll}>Show all</button>}
              {focus.productId !== "" && <span className="qr-focus-note" title="Opened from QR SUMMARY">FROM QR SUMMARY · {focus.productId}</span>}
            </div>

            <div className="qr-table-wrap" ref={virtual.ref} onScroll={virtual.onScroll}>
              {data.columns.length === 0 ? (
                <div className="qr-empty">The query returned no columns.</div>
              ) : selected.length === 0 ? (
                <div className="qr-empty">Select a test date on the left.</div>
              ) : data.rows.length === 0 ? (
                <div className="qr-empty">No data for the selected dates.</div>
              ) : (
                <table className="qr-data-table">
                  <thead>
                    <tr>
                      <th className="qr-expand-head"><span>QR</span><small>group</small></th>
                      {data.layout.map((slot) =>
                        slot === TEST_DATE_INDEX ? (
                          <th key="date" className="qr-test-date-head">
                            <span>Test date</span>
                            <small>{displayCell(data.sortColumnName)}</small>
                            {filterBox(TEST_DATE_INDEX, "test date")}
                          </th>
                        ) : (
                          <th key={slot} title={data.columns[slot].dataType}>
                            <span>{data.columns[slot].name}</span>
                            <small>{data.columns[slot].dataType}</small>
                            {filterBox(slot, data.columns[slot].name)}
                          </th>
                        ),
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {lines.length === 0 && (
                      <tr>
                        <td className="qr-no-match" colSpan={colCount}>
                          No row matches the current filters. Clear a highlighted filter box above, or press Show all.
                        </td>
                      </tr>
                    )}
                    <tr aria-hidden="true"><td colSpan={colCount} style={{ height: virtual.before, padding: 0, border: 0 }} /></tr>
                    {lines.slice(virtual.first, virtual.last).map(renderLine)}
                    <tr aria-hidden="true"><td colSpan={colCount} style={{ height: virtual.after, padding: 0, border: 0 }} /></tr>
                  </tbody>
                </table>
              )}
            </div>
          </section>
        </div>
      )}
    </div>
  );
}
