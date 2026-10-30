import { useEffect, useMemo, useRef, useState } from "react";
import { ngRateApi, type NgRateSetup, type ReportResult } from "../ngRateApi";
import { downloadBlob } from "../toolsApi";
import { ReportSummaryChart } from "./ngRate/charts";
import {
  DATE_LIMIT, MAX_REASONS, MONTH_LIMIT, WEEK_LIMIT, displayReasonRows, toggled, visibleCols, type ColumnSet,
} from "./ngRate/logic";
import { HierCard, ModelGroupPicker, SetupPanel, Spinner, ViewNav, activeProductGroupFor } from "./ngRate/parts";
import { ExpandLinks, NgTable, ProcessTable, ReasonTable, SummaryTable, reasonKey } from "./ngRate/tables";
import { usePpmMaxHighlight } from "./ngRate/tableStyle";
import { useReportRun } from "./ngRate/useReportRun";
import "../ngRate.css";

/** NG Rate — By Model (port of Components/Pages/NgRatePage.razor): one tab per picked group / model. */
export default function NgRate() {
  const page = useReportRun();
  const root = useRef<HTMLDivElement>(null);
  usePpmMaxHighlight(root);

  const [selectedIds, setSelectedIds] = useState<Set<number>>(new Set());
  const [selectedEntries, setSelectedEntries] = useState<Set<string>>(new Set());
  const [activePg, setActivePg] = useState("ETC");

  const [result, setResult] = useState<ReportResult | null>(null);
  const [tab, setTab] = useState(0);
  const [runId, setRunId] = useState(0);
  const [exporting, setExporting] = useState(false);

  const [expandedSummary, setExpandedSummary] = useState<Set<string>>(new Set());
  const [expandedProcess, setExpandedProcess] = useState<Set<number>>(new Set());
  const [expandedNg, setExpandedNg] = useState<Set<number>>(new Set());
  const [expandedReason, setExpandedReason] = useState<Set<string>>(new Set());
  const [collapsedReason, setCollapsedReason] = useState<Set<string>>(new Set());
  // Group rows hidden from every expansion (✕ on a Summary group row); does not change totals or the chart.
  const [hiddenGroups, setHiddenGroups] = useState<Set<string>>(new Set());
  const [dateExpanded, setDateExpanded] = useState(false);
  const [weekExpanded, setWeekExpanded] = useState(false);
  const [monthExpanded, setMonthExpanded] = useState(false);

  const reports = result?.reports ?? [];
  const report = reports[tab]?.report ?? null;

  const cols = useMemo<ColumnSet>(() => ({
    date: report ? visibleCols(report.dateCols, DATE_LIMIT, dateExpanded) : [],
    week: report ? visibleCols(report.weekCols, WEEK_LIMIT, weekExpanded) : [],
    month: report ? visibleCols(report.monthCols, MONTH_LIMIT, monthExpanded) : [],
  }), [report, dateExpanded, weekExpanded, monthExpanded]);

  const reasonRows = useMemo(() => (report ? displayReasonRows(report.reasonRows, MAX_REASONS) : []), [report]);
  const shown = (group: string) => !hiddenGroups.has(group);

  const pruneSelection = (setup: NgRateSetup) => {
    const ids = new Set(setup.groups.map((g) => g.id));
    const live = new Set(setup.groups.flatMap((g) => g.children.map((c) => c.key)));
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
    setTab(0);
    setExpandedSummary(new Set());
    setExpandedProcess(new Set());
    setExpandedNg(new Set());
    setExpandedReason(new Set());
    setCollapsedReason(new Set());
    setHiddenGroups(new Set());
    setDateExpanded(false);
    setWeekExpanded(false);
    setMonthExpanded(false);
  };

  const getReport = async () => {
    if (page.busy || !page.canRun()) return;
    setResult(null);
    resetView();
    const next = await page.run((onProgress) =>
      ngRateApi.report("by-model", page.start, page.end, [...selectedIds], [...selectedEntries], onProgress));
    if (next) {
      setResult(next);
      setRunId((n) => n + 1);
    }
    await refreshGroups();
  };

  const clearReport = () => {
    setResult(null);
    resetView();
  };

  const switchTab = (index: number) => {
    setTab(index);
    setExpandedProcess(new Set());
    setExpandedNg(new Set());
    setExpandedReason(new Set());
    setCollapsedReason(new Set());
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

  const colToggle = (label: string, total: number, limit: number, expanded: boolean, toggle: () => void) =>
    total > limit && (
      <button type="button" className="expand-btn" onClick={toggle}>
        {label} {expanded ? "▼" : `▶ +${total - limit}`}
      </button>
    );

  return (
    <div className="ngr" ref={root}>
      <ViewNav current="by-model" />

      <SetupPanel
        start={page.start} end={page.end} onStart={page.setStart} onEnd={page.setEnd} onToday={page.setToday}
        busy={page.busy} setup={page.setup} savedDbPath={page.savedDbPath} errorMessage={page.errorMessage}
        onClearResult={page.clearResult} logs={page.logs} hasSuccessfulResult={page.savedDbPath !== null || reports.length > 0}
        onClearLogs={page.clearLogs}>
        {page.setup && (
          <ModelGroupPicker
            setup={page.setup} entries="children" selectedIds={selectedIds} selectedEntries={selectedEntries}
            activeProductGroup={activePg} onActiveProductGroup={setActivePg} busy={page.busy} reportError={page.reportError}
            onRefresh={() => void refreshGroups()} onToggleGroup={(id) => setSelectedIds(toggled(selectedIds, id))}
            onToggleEntry={(key) => setSelectedEntries(toggled(selectedEntries, key))} onGetReport={() => void getReport()} />
        )}
      </SetupPanel>

      {report && (
        <div className="mt-4">
          <div className="d-flex align-items-end justify-content-between mb-0">
            {reports.length > 1 ? (
              <ul className="ngr-tabs mb-0">
                {reports.map((r, i) => (
                  <li key={i}>
                    <button type="button" className={"ngr-tab" + (i === tab ? " active fw-semibold" : "")} onClick={() => switchTab(i)}>
                      {r.label}
                    </button>
                  </li>
                ))}
              </ul>
            ) : <div />}
            <button type="button" className="ngr-btn ngr-btn-outline-success ngr-btn-sm mb-1" style={{ fontSize: 12 }}
              onClick={() => void exportExcel()} disabled={exporting}>
              {exporting ? <Spinner /> : <span>⬇ Export Excel</span>}
            </button>
          </div>
          <div style={{ borderTop: "2px solid #e2e8f0", marginBottom: 12 }} />

          {result && result.hier.length > 0 && (
            <HierCard key={runId} hier={result.hier} cols={cols}
              title="📊 Group → Model Name → Sub Group → Shift Group — TOTAL PPM (collapsible)"
              label="Group / Model Name / Sub Group / Shift Group" />
          )}

          <div className="d-flex align-items-center gap-2 mb-3 border-bottom pb-2">
            <span style={{ fontSize: 18 }}>📊</span>
            <h6 className="mb-0 fw-semibold">Report</h6>
            <span className="small text-muted ms-1">
              {report.dbPath} · {report.totalRows.toLocaleString("en-US")} rows · {report.generatedAt}
            </span>
            {report.modelFilter.length > 0 && (
              <span className="ngr-badge bg-secondary ms-1" style={{ fontSize: 10 }}>
                {report.modelFilter.length} model{report.modelFilter.length === 1 ? "" : "s"} filtered
              </span>
            )}
            <button type="button" className="ngr-link ms-auto text-secondary" onClick={clearReport}>✕ Clear</button>
          </div>

          <div className="ngr-card mb-3 ngrate-report-card">
            <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
              <div className="d-flex align-items-center gap-2 flex-wrap">
                <span>Summary — NG PPM by Process Type</span>
                {hiddenGroups.size > 0 && (
                  <button type="button" className="group-chip-clear ms-2" title="Restore all hidden group rows" onClick={() => setHiddenGroups(new Set())}>
                    ✕ Clear hidden ({hiddenGroups.size})
                  </button>
                )}
              </div>
              <div className="d-flex gap-2 align-items-center">
                {Object.keys(report.groupSummary).length > 0 && (
                  <>
                    <ExpandLinks onExpand={() => setExpandedSummary(new Set(report.summary.map((r) => r.processType)))}
                      onCollapse={() => setExpandedSummary(new Set())} />
                    <span className="text-muted" style={{ fontSize: 10 }}>|</span>
                  </>
                )}
                {colToggle("Date", report.dateCols.length, DATE_LIMIT, dateExpanded, () => setDateExpanded(!dateExpanded))}
                {colToggle("Week", report.weekCols.length, WEEK_LIMIT, weekExpanded, () => setWeekExpanded(!weekExpanded))}
                {colToggle("Month", report.monthCols.length, MONTH_LIMIT, monthExpanded, () => setMonthExpanded(!monthExpanded))}
              </div>
            </div>
            <div className="card-body ngrate-card-body p-2">
              <SummaryTable report={report} cols={cols} expanded={expandedSummary} shown={shown}
                onToggle={(pt) => setExpandedSummary(toggled(expandedSummary, pt))}
                groupLead={(grp) => (
                  <td className="row-hide-cell">
                    <button type="button" className="row-hide-btn" title="Hide this group row (Clear to restore)"
                      onClick={() => setHiddenGroups(new Set(hiddenGroups).add(grp))}>✕</button>
                  </td>
                )} />
              <div style={{ position: "relative", height: 200, marginTop: 12 }}>
                <ReportSummaryChart summary={report.summary} cols={cols} />
              </div>
            </div>
          </div>

          <div className="ngr-card mb-3 ngrate-report-card">
            <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
              <span>Process 10 NG — Top 10 Worst Processes by PPM</span>
              {report.top10Process.some((r) => r.groups.length > 0) && (
                <ExpandLinks onExpand={() => setExpandedProcess(new Set(report.top10Process.map((r) => r.rank)))}
                  onCollapse={() => setExpandedProcess(new Set())} />
              )}
            </div>
            <div className="card-body ngrate-card-body p-2">
              <ProcessTable rows={report.top10Process} cols={cols} expanded={expandedProcess} shown={shown}
                onToggle={(rank) => setExpandedProcess(toggled(expandedProcess, rank))} />
            </div>
          </div>

          <div className="ngr-card mb-3 ngrate-report-card">
            <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
              <span>Worst 10 NG — Top 10 Worst NG Names by PPM</span>
              {report.top10Ng.some((r) => r.groups.length > 0) && (
                <ExpandLinks onExpand={() => setExpandedNg(new Set(report.top10Ng.map((r) => r.rank)))}
                  onCollapse={() => setExpandedNg(new Set())} />
              )}
            </div>
            <div className="card-body ngrate-card-body p-2">
              <NgTable rows={report.top10Ng} cols={cols} expanded={expandedNg} shown={shown}
                onToggle={(rank) => setExpandedNg(toggled(expandedNg, rank))} />
            </div>
          </div>

          {report.reasonRows.length > 0 && (
            <div className="ngr-card mb-3 ngrate-report-card">
              <div className="card-header ngrate-card-header small fw-semibold d-flex justify-content-between align-items-center">
                <span>Reason — Defect Cause Classification</span>
                {report.reasonRows.some((r) => !r.isTotal && r.groups.length > 0) && (
                  <ExpandLinks onExpand={() => setExpandedReason(new Set(reasonRows.filter((r) => !r.isTotal).map(reasonKey)))}
                    onCollapse={() => setExpandedReason(new Set())} />
                )}
              </div>
              <div className="card-body ngrate-card-body p-2">
                <ReasonTable rows={reasonRows} cols={cols} expanded={expandedReason} shown={shown}
                  onToggle={(key) => setExpandedReason(toggled(expandedReason, key))}
                  collapsed={collapsedReason} onToggleCategory={(reason) => setCollapsedReason(toggled(collapsedReason, reason))} />
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
