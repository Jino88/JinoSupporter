import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import { ngRateApi, type NgRateSetup, type Report, type ReportResult } from "../ngRateApi";
import { downloadBlob } from "../toolsApi";
import { ReportSummaryChart } from "./ngRate/charts";
import { DATE_LIMIT, MONTH_LIMIT, WEEK_LIMIT, displayReasonRows, toggled, visibleCols, type ColumnSet } from "./ngRate/logic";
import { HierCard, ModelGroupPicker, SetupPanel, Spinner, ViewNav, activeProductGroupFor } from "./ngRate/parts";
import { ExpandLinks, NgTable, ProcessTable, ReasonTable, SummaryTable, reasonKey } from "./ngRate/tables";
import { usePpmMaxHighlight } from "./ngRate/tableStyle";
import { useReportRun } from "./ngRate/useReportRun";
import "../ngRate.css";

const SECTIONS = [
  { key: "summary", label: "Summary" },
  { key: "process", label: "Process 10 NG" },
  { key: "ng", label: "Worst 10 NG" },
  { key: "reason", label: "Reason" },
] as const;
type Section = (typeof SECTIONS)[number]["key"];

/** Per-model expand state: the page keeps one set per report index. */
type PerModel<T> = Record<number, Set<T>>;
const setOf = <T,>(state: PerModel<T>, mi: number): Set<T> => state[mi] ?? new Set<T>();

/**
 * NG Rate — By Section (port of Components/Pages/NgRateAllPage.razor): every picked group / model is
 * stacked vertically, one defect-report section (Summary / Process / NG / Reason) at a time.
 */
export default function NgRateAll() {
  const page = useReportRun();
  const root = useRef<HTMLDivElement>(null);
  usePpmMaxHighlight(root);

  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [selectedEntries, setSelectedEntries] = useState<Set<string>>(new Set());
  const [activePg, setActivePg] = useState("ETC");

  const [result, setResult] = useState<ReportResult | null>(null);
  const [runId, setRunId] = useState(0);
  const [section, setSection] = useState<Section>("summary");
  const [exporting, setExporting] = useState(false);

  const [expandedSummary, setExpandedSummary] = useState<PerModel<string>>({});
  const [expandedProcess, setExpandedProcess] = useState<PerModel<number>>({});
  const [expandedNg, setExpandedNg] = useState<PerModel<number>>({});
  const [expandedReason, setExpandedReason] = useState<PerModel<string>>({});
  const [collapsedReason, setCollapsedReason] = useState<PerModel<string>>({});
  // Group chips switched off in a model's Summary header (their ↳ rows are hidden).
  const [hiddenSummaryGroups, setHiddenSummaryGroups] = useState<PerModel<string>>({});
  const [dateExpanded, setDateExpanded] = useState(false);
  const [weekExpanded, setWeekExpanded] = useState(false);
  const [monthExpanded, setMonthExpanded] = useState(false);

  const reports = useMemo(() => result?.reports ?? [], [result]);

  // One column set per report, memoised so the charts only redraw when the visible columns change.
  const allCols = useMemo<ColumnSet[]>(() => reports.map(({ report: r }) => ({
    date: visibleCols(r.dateCols, DATE_LIMIT, dateExpanded),
    week: visibleCols(r.weekCols, WEEK_LIMIT, weekExpanded),
    month: visibleCols(r.monthCols, MONTH_LIMIT, monthExpanded),
  })), [reports, dateExpanded, weekExpanded, monthExpanded]);

  const pruneSelection = (setup: NgRateSetup) => {
    const ids = new Set(setup.groups.map((g) => g.id));
    const live = new Set(setup.groups.flatMap((g) => g.midChildren.map((c) => c.key)));
    setSelectedIds((cur) => new Set([...cur].filter((id) => ids.has(id))));
    setSelectedEntries((cur) => new Set([...cur].filter((k) => live.has(k))));
  };

  // Like the Blazor RefreshGroups: stay on the product group while it has groups, else the first that has.
  useEffect(() => {
    if (page.setup) setActivePg((cur) => activeProductGroupFor(page.setup!, cur));
  }, [page.setup]);

  const refreshGroups = async () => {
    const setup = await page.loadSetup();
    if (setup) pruneSelection(setup);
  };

  const resetView = () => {
    setSection("summary");
    setExpandedSummary({});
    setExpandedProcess({});
    setExpandedNg({});
    setExpandedReason({});
    setCollapsedReason({});
    setHiddenSummaryGroups({});
    setDateExpanded(false);
    setWeekExpanded(false);
    setMonthExpanded(false);
  };

  const getReport = async () => {
    if (page.busy || !page.canRun()) return;
    setResult(null);
    resetView();
    const next = await page.run((onProgress) =>
      ngRateApi.report("by-section", page.start, page.end, [...selectedIds], [...selectedEntries], onProgress));
    if (next) {
      setResult(next);
      setRunId((n) => n + 1);
    }
    await refreshGroups();
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

  function update<T>(setter: (fn: (s: PerModel<T>) => PerModel<T>) => void, mi: number, next: (current: Set<T>) => Set<T>) {
    setter((state) => ({ ...state, [mi]: next(setOf(state, mi)) }));
  }

  const maxOf = (pick: (r: Report) => number) => Math.max(0, ...reports.map((r) => pick(r.report)));
  const colToggle = (label: string, max: number, limit: number, expanded: boolean, toggle: () => void) =>
    max > limit && (
      <button type="button" className="expand-btn" onClick={toggle}>
        {label} {expanded ? "Collapse" : `+${max - limit}`}
      </button>
    );

  return (
    <div className="ngr" ref={root}>
      <ViewNav current="by-section" />

      <SetupPanel
        start={page.start} end={page.end} onStart={page.setStart} onEnd={page.setEnd} onToday={page.setToday}
        busy={page.busy} setup={page.setup} savedDbPath={page.savedDbPath} errorMessage={page.errorMessage}
        onClearResult={page.clearResult} logs={page.logs} hasSuccessfulResult={page.savedDbPath !== null || reports.length > 0}
        onClearLogs={page.clearLogs}>
        {page.setup && (
          <ModelGroupPicker
            setup={page.setup} entries="midChildren" selectedIds={selectedIds} selectedEntries={selectedEntries}
            activeProductGroup={activePg} onActiveProductGroup={setActivePg} busy={page.busy} reportError={page.reportError}
            onRefresh={() => void refreshGroups()} onToggleGroup={(id) => setSelectedIds(toggled(selectedIds, id))}
            onToggleEntry={(key) => setSelectedEntries(toggled(selectedEntries, key))} onGetReport={() => void getReport()} />
        )}
      </SetupPanel>

      {reports.length > 0 && result && (
        <div className="mt-4">
          <div className="d-flex align-items-end justify-content-between mb-0 flex-wrap gap-2">
            <ul className="ngr-tabs mb-0">
              {SECTIONS.map((s) => (
                <li key={s.key}>
                  <button type="button" className={"ngr-tab" + (section === s.key ? " active fw-semibold" : "")} onClick={() => setSection(s.key)}>
                    {s.label}
                  </button>
                </li>
              ))}
            </ul>
            <div className="d-flex gap-2 align-items-center">
              {colToggle("Date", maxOf((r) => r.dateCols.length), DATE_LIMIT, dateExpanded, () => setDateExpanded(!dateExpanded))}
              {colToggle("Week", maxOf((r) => r.weekCols.length), WEEK_LIMIT, weekExpanded, () => setWeekExpanded(!weekExpanded))}
              {colToggle("Month", maxOf((r) => r.monthCols.length), MONTH_LIMIT, monthExpanded, () => setMonthExpanded(!monthExpanded))}
              <button type="button" className="ngr-btn ngr-btn-outline-success ngr-btn-sm" style={{ fontSize: 12 }}
                onClick={() => void exportExcel()} disabled={exporting}>
                {exporting ? <Spinner /> : <span>Export Excel</span>}
              </button>
              <button type="button" className="ngr-link text-secondary" onClick={() => { setResult(null); resetView(); }}>Clear</button>
            </div>
          </div>
          <div style={{ borderTop: "2px solid #e2e8f0", marginBottom: 12 }} />

          {result.hier.length > 0 && (
            <HierCard key={runId} hier={result.hier} cols={allCols[0]}
              title="Group -> Model -> Line -> Shift - TOTAL PPM (collapsible)" label="Group / Model / Line / Shift" />
          )}

          {reports.map(({ label, report: rpt }, mi) => {
            const cols = allCols[mi];
            const hidden = setOf(hiddenSummaryGroups, mi);
            const shownInSummary = (grp: string) => !hidden.has(grp);
            return (
              <Fragment key={mi}>
                <div className={"d-flex align-items-center gap-2 mb-2 " + (mi > 0 ? "mt-4 pt-3 border-top" : "mt-2")}>
                  <span className="ngr-badge bg-secondary" style={{ fontSize: 11 }}>{mi + 1}</span>
                  <h6 className="mb-0 fw-semibold" style={{ color: "#1e40af" }}>{label}</h6>
                  <span className="small text-muted ms-1">
                    {rpt.dbPath} | {rpt.totalRows.toLocaleString("en-US")} rows | {rpt.generatedAt}
                  </span>
                </div>

                {section === "summary" && (
                  <div className="ngr-card mb-3 ngrate-report-card">
                    <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
                      <div className="d-flex align-items-center gap-2 flex-wrap">
                        <span>Summary - NG PPM by Process Type</span>
                        {rpt.groupNames.length > 0 && (
                          <div className="d-flex flex-wrap gap-1 ms-1">
                            {rpt.groupNames.map((grp) => (
                              <span key={grp} className={"group-chip" + (shownInSummary(grp) ? " active" : "")}
                                title="Toggle this group's rows in the summary table"
                                onClick={() => update(setHiddenSummaryGroups, mi, (s) => toggled(s, grp))}>{grp}</span>
                            ))}
                          </div>
                        )}
                      </div>
                      {Object.keys(rpt.groupSummary).length > 0 && (
                        <ExpandLinks onExpand={() => update(setExpandedSummary, mi, () => new Set(rpt.summary.map((r) => r.processType)))}
                          onCollapse={() => update(setExpandedSummary, mi, () => new Set())} />
                      )}
                    </div>
                    <div className="card-body ngrate-card-body p-2">
                      <SummaryTable report={rpt} cols={cols} expanded={setOf(expandedSummary, mi)} shown={shownInSummary}
                        onToggle={(pt) => update(setExpandedSummary, mi, (s) => toggled(s, pt))} />
                      <div style={{ position: "relative", height: 200, marginTop: 12 }}>
                        <ReportSummaryChart summary={rpt.summary} cols={cols} />
                      </div>
                    </div>
                  </div>
                )}

                {section === "process" && (
                  <div className="ngr-card mb-3 ngrate-report-card">
                    <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
                      <span>Process 10 NG - Top 10 Worst Processes by PPM</span>
                      {rpt.top10Process.some((r) => r.groups.length > 0) && (
                        <ExpandLinks onExpand={() => update(setExpandedProcess, mi, () => new Set(rpt.top10Process.map((r) => r.rank)))}
                          onCollapse={() => update(setExpandedProcess, mi, () => new Set())} />
                      )}
                    </div>
                    <div className="card-body ngrate-card-body p-2">
                      <ProcessTable rows={rpt.top10Process} cols={cols} expanded={setOf(expandedProcess, mi)} shown={() => true}
                        onToggle={(rank) => update(setExpandedProcess, mi, (s) => toggled(s, rank))} />
                    </div>
                  </div>
                )}

                {section === "ng" && (
                  <div className="ngr-card mb-3 ngrate-report-card">
                    <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
                      <span>Worst 10 NG - Top 10 Worst NG Names by PPM</span>
                      {rpt.top10Ng.some((r) => r.groups.length > 0) && (
                        <ExpandLinks onExpand={() => update(setExpandedNg, mi, () => new Set(rpt.top10Ng.map((r) => r.rank)))}
                          onCollapse={() => update(setExpandedNg, mi, () => new Set())} />
                      )}
                    </div>
                    <div className="card-body ngrate-card-body p-2">
                      <NgTable rows={rpt.top10Ng} cols={cols} expanded={setOf(expandedNg, mi)} shown={() => true}
                        onToggle={(rank) => update(setExpandedNg, mi, (s) => toggled(s, rank))} />
                    </div>
                  </div>
                )}

                {section === "reason" && rpt.reasonRows.length > 0 && (
                  <div className="ngr-card mb-3 ngrate-report-card">
                    <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
                      <span>Reason - Defect Cause Classification</span>
                      {rpt.reasonRows.some((r) => !r.isTotal && r.groups.length > 0) && (
                        <ExpandLinks
                          onExpand={() => update(setExpandedReason, mi, () => new Set(rpt.reasonRows.filter((r) => !r.isTotal).map(reasonKey)))}
                          onCollapse={() => update(setExpandedReason, mi, () => new Set())} />
                      )}
                    </div>
                    <div className="card-body ngrate-card-body p-2">
                      <ReasonTable rows={displayReasonRows(rpt.reasonRows)} cols={cols} expanded={setOf(expandedReason, mi)} shown={() => true}
                        onToggle={(key) => update(setExpandedReason, mi, (s) => toggled(s, key))}
                        collapsed={setOf(collapsedReason, mi)} onToggleCategory={(reason) => update(setCollapsedReason, mi, (s) => toggled(s, reason))} />
                    </div>
                  </div>
                )}
              </Fragment>
            );
          })}
        </div>
      )}
    </div>
  );
}
