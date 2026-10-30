import { Fragment, useMemo, useRef, useState } from "react";
import { ngRateApi, type ByGroupResult, type Period } from "../ngRateApi";
import { downloadBlob } from "../toolsApi";
import { GroupChart } from "./ngRate/charts";
import {
  BY_GROUP_DATE_LIMIT, MONTH_LIMIT, WEEK_LIMIT, allKeys, chartAxis, formatPpm, hasAnyValue, meanOfPositive, ppmClass, ppmOf,
  processTypeClass, toggled, visibleCols, type ColumnSet,
} from "./ngRate/logic";
import { BlockPeriodCells, BlockPeriodHeads, HierSubRows, SetupPanel, SimpleGroupPicker, Spinner, ViewNav } from "./ngRate/parts";
import { usePpmMaxHighlight } from "./ngRate/tableStyle";
import { useReportRun } from "./ngRate/useReportRun";
import "../ngRate.css";

const GROUP_SUB_CLASSES = { sub1: "grp-sub-row grp-sub1-row", sub2: "grp-sub-row grp-sub2-row", lineShift: "grp-ls-row", expandable: "grp-sub-expandable" };

/**
 * NG Rate — By Group (port of Components/Pages/NgRateByGroupPage.razor): picked Model Groups × period,
 * Group → Process Type and Group → Model → Sub → Shift summaries, trend charts and a Top-10 per group.
 */
export default function NgRateByGroup() {
  const page = useReportRun();
  const root = useRef<HTMLDivElement>(null);
  usePpmMaxHighlight(root);

  const [selected, setSelected] = useState<string[]>([]);
  const [result, setResult] = useState<ByGroupResult | null>(null);
  const [exporting, setExporting] = useState(false);
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set());
  const [expandedGroupsModel, setExpandedGroupsModel] = useState<Set<string>>(new Set());
  const [expandedMids, setExpandedMids] = useState<Set<string>>(new Set());
  const [expandedSubs, setExpandedSubs] = useState<Set<string>>(new Set());
  const [dateExpanded, setDateExpanded] = useState(false);
  const [weekExpanded, setWeekExpanded] = useState(false);
  const [monthExpanded, setMonthExpanded] = useState(false);

  const groups = page.setup?.groups ?? [];

  const cols = useMemo<ColumnSet>(() => ({
    date: result ? visibleCols(result.dateCols, BY_GROUP_DATE_LIMIT, dateExpanded) : [],
    week: result ? visibleCols(result.weekCols, WEEK_LIMIT, weekExpanded) : [],
    month: result ? visibleCols(result.monthCols, MONTH_LIMIT, monthExpanded) : [],
  }), [result, dateExpanded, weekExpanded, monthExpanded]);
  const keys = allKeys(cols);
  const totalColCount = keys.length;

  const overall = (key: string) => meanOfPositive((result?.groups ?? []).map((g) => ppmOf(g.total, key)));

  const chart = useMemo(() => {
    const axis = chartAxis(cols);
    return { labels: axis.labels, series: (result?.groups ?? []).map((g) => ({ name: g.name, values: axis.values(g.total) })) };
  }, [cols, result]);

  const reloadGroups = async () => {
    const setup = await page.loadSetup();
    if (setup) {
      const names = new Set(setup.groups.map((g) => g.name));
      setSelected((cur) => cur.filter((n) => names.has(n)));
    }
  };

  const getReport = async () => {
    if (page.busy || selected.length === 0 || !page.canRun()) return;
    setResult(null);
    const next = await page.run((onProgress) => ngRateApi.byGroup(page.start, page.end, selected, onProgress));
    if (next) setResult(next);
    await reloadGroups();
  };

  const exportExcel = async () => {
    if (exporting || !result?.exportId) return;
    setExporting(true);
    page.setReportError("");
    try {
      const file = await ngRateApi.exportExcel(result.exportId);
      downloadBlob(file.blob, file.fileName);
    } catch (e) {
      page.setReportError(`Export failed: ${(e as Error).message}`);
    } finally {
      setExporting(false);
    }
  };

  const toggleButton = (variant: string, label: string, all: Period[], cap: number, expanded: boolean, toggle: () => void) => {
    const hidden = Math.max(0, all.length - Math.min(all.length, cap));
    if (!expanded && hidden === 0) return null;
    return (
      <button type="button" className={"ngr-btn ngr-btn-outline-" + variant} onClick={toggle}>
        {expanded ? `◂ ${label} (${all.length - cap} hide)` : `▸ ${label} (+${hidden})`}
      </button>
    );
  };

  const totalRow = (
    <tr className="grp-overall-row">
      <td className="label-td">
        <strong>TOTAL</strong>
        <span className="text-muted ms-1" style={{ fontSize: 10 }}>{result?.groups.length ?? 0} groups avg</span>
      </td>
      <BlockPeriodCells cols={cols} value={overall} />
    </tr>
  );

  const noData = <div className="text-muted small p-4 text-center">No data.</div>;

  return (
    <div className="ngr" ref={root}>
      <ViewNav current="by-group" />

      <SetupPanel flow
        start={page.start} end={page.end} onStart={page.setStart} onEnd={page.setEnd} onToday={page.setToday}
        busy={page.busy} setup={page.setup} savedDbPath={page.savedDbPath} errorMessage={page.errorMessage}
        showClearResult={false} logs={page.logs} hasSuccessfulResult={result !== null} onClearLogs={page.clearLogs}>
        <SimpleGroupPicker
          groups={groups} selected={selected} busy={page.busy} credentialsConfigured={page.setup?.credentialsConfigured ?? false}
          onReload={() => void reloadGroups()} onSelectAll={() => setSelected(groups.map((g) => g.name))} onClear={() => setSelected([])}
          onToggle={(name, checked) => setSelected((cur) => (checked ? (cur.includes(name) ? cur : [...cur, name]) : cur.filter((n) => n !== name)))}
          onGetReport={() => void getReport()} />
      </SetupPanel>

      {/* Summary (per group × period) */}
      <div className="ngr-card mt-3 ngrate-report-card">
        <div className="card-header ngrate-card-header small fw-semibold d-flex align-items-center gap-2">
          <span>📊 Summary — group × Date / Week / Month (PPM)</span>
          {result && (
            <>
              <span className="text-muted" style={{ fontSize: 11 }}>
                D {cols.date.length}/{result.dateCols.length} · W {cols.week.length}/{result.weekCols.length} · M {cols.month.length}/{result.monthCols.length} · {result.totalRows} rows
              </span>
              <div className="ngr-btn-group ngr-btn-group-sm ms-auto">
                {toggleButton("primary", "Date", result.dateCols, BY_GROUP_DATE_LIMIT, dateExpanded, () => setDateExpanded(!dateExpanded))}
                {toggleButton("success", "Week", result.weekCols, WEEK_LIMIT, weekExpanded, () => setWeekExpanded(!weekExpanded))}
                {toggleButton("warning", "Month", result.monthCols, MONTH_LIMIT, monthExpanded, () => setMonthExpanded(!monthExpanded))}
              </div>
              <button type="button" className="ngr-btn ngr-btn-outline-success ngr-btn-sm" onClick={() => void exportExcel()} disabled={exporting}>
                {exporting ? <><Spinner /> <span>Exporting...</span></> : <span>Export Excel</span>}
              </button>
            </>
          )}
        </div>
        <div className="card-body ngrate-card-body p-0">
          {page.reportError ? (
            <div className="ngr-alert ngr-alert-danger m-3 small">{page.reportError}</div>
          ) : !result ? (
            <div className="text-muted small p-4 text-center">Select groups, then click <strong>Generate Report</strong>.</div>
          ) : totalColCount === 0 ? noData : (
            <div style={{ overflowX: "auto" }}>
              <table className="pivot-table text-fit-table">
                <BlockPeriodHeads cols={cols} label="Group" labelStyle={{ width: 220 }} />
                <tbody>
                  {totalRow}
                  {result.groups.map((g) => {
                    const open = expandedGroups.has(g.name);
                    return (
                      <Fragment key={g.name}>
                        <tr className="grp-total-row" onClick={() => setExpandedGroups(toggled(expandedGroups, g.name))}>
                          <td className="label-td">
                            <span className="tog-icon">{open ? "▼" : "▶"}</span>
                            <span className="ngrate-hier-level">GROUP</span>
                            <strong>{g.name}</strong>
                          </td>
                          <BlockPeriodCells cols={cols} value={(k) => ppmOf(g.total, k)} />
                        </tr>
                        {open && g.processTypes.map((pt, i) => (
                          <tr key={i} className="grp-sub-row">
                            <td className="label-td ps-4"><span className={"pt-badge " + processTypeClass(pt.processType)}>{pt.processType}</span></td>
                            <BlockPeriodCells cols={cols} value={(k) => ppmOf(pt.ppm, k)} />
                          </tr>
                        ))}
                      </Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      {/* Chart (per-group PPM trend) */}
      {result && totalColCount > 0 && result.groups.length > 0 && (
        <div className="ngr-card mt-3 ngrate-report-card">
          <div className="card-header ngrate-card-header small fw-semibold d-flex align-items-center gap-2">
            <span>📈 Group trend — PPM (Date / Week / Month)</span>
          </div>
          <div className="card-body ngrate-card-body" style={{ height: 320 }}>
            <GroupChart labels={chart.labels} series={chart.series} />
          </div>
        </div>
      )}

      {/* Summary 2 (per group × model) */}
      <div className="ngr-card mt-3 ngrate-report-card">
        <div className="card-header ngrate-card-header small fw-semibold d-flex align-items-center gap-2">
          <span>📊 Summary — Group → Model Name × Date / Week / Month (PPM)</span>
          {result && <span className="text-muted" style={{ fontSize: 11 }}>Model Name PPM = same combo-merged as per-model</span>}
        </div>
        <div className="card-body ngrate-card-body p-0">
          {!result ? (
            <div className="text-muted small p-4 text-center">Displayed after Generate Report.</div>
          ) : totalColCount === 0 ? noData : (
            <div style={{ overflowX: "auto" }}>
              <table className="pivot-table text-fit-table">
                <BlockPeriodHeads cols={cols} label="Group / Model" labelStyle={{ width: 260 }} />
                <tbody>
                  {totalRow}
                  {result.groups.map((g) => {
                    const open = expandedGroupsModel.has(g.name);
                    const mids = g.mids.filter((m) => hasAnyValue(m.ppm, keys));
                    return (
                      <Fragment key={g.name}>
                        <tr className="grp-total-row" onClick={() => setExpandedGroupsModel(toggled(expandedGroupsModel, g.name))}>
                          <td className="label-td">
                            <span className="tog-icon">{open ? "▼" : "▶"}</span>
                            <span className="ngrate-hier-level">GROUP</span>
                            <strong>{g.name}</strong>
                            <span className="text-muted ms-1" style={{ fontSize: 10 }}>{mids.length} models</span>
                          </td>
                          <BlockPeriodCells cols={cols} value={(k) => ppmOf(g.total, k)} />
                        </tr>
                        {open && mids.map((m) => {
                          const midKey = `${g.name}::${m.material}`;
                          const midOpen = expandedMids.has(midKey);
                          return (
                            <Fragment key={midKey}>
                              <tr className="grp-sub-row grp-mid-row" onClick={() => setExpandedMids(toggled(expandedMids, midKey))}>
                                <td className="label-td ps-4" style={{ fontSize: 11 }}>
                                  <span className="tog-icon">{midOpen ? "▼" : "▶"}</span>
                                  <span className="ngrate-hier-level">MODEL</span>
                                  <span className="ms-1">{m.material}</span>
                                  <span className="text-muted ms-1" style={{ fontSize: 10 }}>{m.lineShiftCount}</span>
                                </td>
                                <BlockPeriodCells cols={cols} value={(k) => ppmOf(m.ppm, k)} />
                              </tr>
                              {midOpen && (
                                <HierSubRows mid={m} cols={cols} expanded={expandedSubs} onToggle={(id) => setExpandedSubs(toggled(expandedSubs, id))}
                                  classes={GROUP_SUB_CLASSES} basePad={24} />
                              )}
                            </Fragment>
                          );
                        })}
                      </Fragment>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      {/* Chart 2 (same trend as above, as on the Blazor page) */}
      {result && totalColCount > 0 && result.groups.length > 0 && (
        <div className="ngr-card mt-3 ngrate-report-card">
          <div className="card-header ngrate-card-header small fw-semibold d-flex align-items-center gap-2">
            <span>📈 Group trend — PPM (Date / Week / Month)</span>
          </div>
          <div className="card-body ngrate-card-body" style={{ height: 320 }}>
            <GroupChart labels={chart.labels} series={chart.series} />
          </div>
        </div>
      )}

      {/* Top-10 per group (Model × Process × NG) */}
      {result && result.groups.length > 0 && (
        <div className="ngr-card mt-3 ngrate-report-card">
          <div className="card-header ngrate-card-header small fw-semibold d-flex align-items-center gap-2">
            <span>🥇 Top 10 per group — Model Name × Process × NG</span>
            <span className="text-muted" style={{ fontSize: 11 }}>
              {result.sortRefKey
                ? `Sorted: previous-week PPM, descending (ref = ${result.sortRefBucket} ${result.sortRefLabel})`
                : "No sort reference available."}
            </span>
          </div>
          <div className="card-body ngrate-card-body p-0">
            {result.groups.map((g) => (
              <Fragment key={g.name}>
                <div className="px-3 py-2 top10-group-head">
                  <span className="ngrate-hier-level">GROUP</span>
                  <strong className="me-2">{g.name}</strong>
                  <span className="text-muted small">{g.top10.length} rows</span>
                </div>
                {g.top10.length === 0 ? (
                  <div className="text-muted small px-3 pb-3">(no data)</div>
                ) : (
                  <Top10Table result={result} rows={g.top10} />
                )}
              </Fragment>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function Top10Table({ result, rows }: { result: ByGroupResult; rows: ByGroupResult["groups"][number]["top10"] }) {
  const { top10DateCols: dates, top10WeekCols: weeks, top10MonthCols: months } = result;
  const cell = (ppm: Record<string, number>, key: string) => {
    const v = ppmOf(ppm, key);
    return <td key={key} className={ppmClass(v)}>{formatPpm(v)}</td>;
  };
  return (
    <div style={{ overflowX: "auto" }}>
      <table className="pivot-table top10-table">
        <colgroup>
          <col style={{ width: 48 }} />
          <col style={{ width: 190 }} />
          <col style={{ width: 360 }} />
          <col style={{ width: 190 }} />
          {dates.map((c) => <col key={c.key} style={{ width: 70 }} />)}
          {dates.length > 0 && <col style={{ width: 8 }} />}
          {weeks.map((c) => <col key={c.key} style={{ width: 70 }} />)}
          {weeks.length > 0 && <col style={{ width: 8 }} />}
          {months.map((c) => <col key={c.key} style={{ width: 70 }} />)}
        </colgroup>
        <thead>
          <tr>
            <th className="label-th" rowSpan={2}>#</th>
            <th className="label-th" rowSpan={2}>Model Name</th>
            <th className="label-th process-name-col" rowSpan={2}>Process</th>
            <th className="label-th ng-name-col" rowSpan={2}>NG Name</th>
            {dates.length > 0 && <><th className="blk-date" colSpan={dates.length}>Date</th><th className="sep-th" rowSpan={2} /></>}
            {weeks.length > 0 && <><th className="blk-week" colSpan={weeks.length}>Week</th><th className="sep-th" rowSpan={2} /></>}
            {months.length > 0 && <th className="blk-month" colSpan={months.length}>Month</th>}
          </tr>
          <tr>
            {dates.map((c) => <th key={c.key}>{c.header}</th>)}
            {weeks.map((c) => <th key={c.key}>{c.header}</th>)}
            {months.map((c) => <th key={c.key}>{c.header}</th>)}
          </tr>
        </thead>
        <tbody>
          {rows.map((d, i) => (
            <tr key={i}>
              <td className="label-td n text-muted">{i + 1}</td>
              <td className="label-td font-monospace" style={{ fontSize: 11 }}>{d.material}</td>
              <td className="label-td process-name-col" style={{ fontSize: 11 }}>
                <span className="process-cell">
                  <span className={"pt-badge " + processTypeClass(d.processType)}>{d.processType}</span>
                  <span className="process-label">{d.processName}</span>
                </span>
              </td>
              <td className="label-td ng-name-col" style={{ fontSize: 11 }}>{d.ngName}</td>
              {dates.map((c) => cell(d.datePpm, c.key))}
              {dates.length > 0 && <td className="sep-td" />}
              {weeks.map((c) => cell(d.weekPpm, c.key))}
              {weeks.length > 0 && <td className="sep-td" />}
              {months.map((c) => cell(d.monthPpm, c.key))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
