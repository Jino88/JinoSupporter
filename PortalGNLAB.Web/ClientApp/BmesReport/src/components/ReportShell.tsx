import { useEffect, useMemo, useRef, useState, type KeyboardEvent } from "react";
import { TAB_KEYS, type BmesReportDocument, type TabKey } from "../contract";
import { loadStoredPreferences, selectInitialTab, storePreferences, type ViewerPreferences } from "../logic";
import { ReportTabContent } from "./ReportTabContent";
import { ViewerToolbar } from "./ViewerToolbar";

const LABELS: Record<TabKey, string> = {
  daily: "Daily",
  weekly: "Weekly",
  kpi: "KPI",
  "cause-monthly": "원인 비중",
  fcost: "F-COST",
  "fcost-all": "F-COST 전체",
  "fcost-weekly": "목표 불량률",
  "fcost-weekly-all": "목표 불량률 전체",
};

export function ReportShell({ report }: { report: BmesReportDocument }) {
  const [activeTab, setActiveTab] = useState<TabKey>(() => selectInitialTab(report.viewerDefaults.defaultTab));
  const [preferences, setPreferences] = useState<ViewerPreferences>(() => loadStoredPreferences(report.viewerDefaults));
  const tabRefs = useRef<Array<HTMLButtonElement | null>>([]);
  const selectedGroups = useMemo(() => report.request.groups.map((group) => group.name).join(", "), [report.request.groups]);

  useEffect(() => storePreferences(preferences), [preferences]);

  const activate = (index: number) => {
    const wrapped = (index + TAB_KEYS.length) % TAB_KEYS.length;
    setActiveTab(TAB_KEYS[wrapped]);
    requestAnimationFrame(() => tabRefs.current[wrapped]?.focus());
  };

  const handleTabKey = (event: KeyboardEvent<HTMLButtonElement>, index: number) => {
    if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      event.preventDefault(); activate(index + 1);
    } else if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
      event.preventDefault(); activate(index - 1);
    } else if (event.key === "Home") {
      event.preventDefault(); activate(0);
    } else if (event.key === "End") {
      event.preventDefault(); activate(TAB_KEYS.length - 1);
    }
  };

  return (
    <main className="report-shell">
      <header className="report-header">
        <div>
          <p className="eyebrow">BMES INTEGRATED REPORT</p>
          <h1>BMES 통합 리포트</h1>
          <p className="report-context">
            <time dateTime={report.request.startDate}>{report.request.startDate}</time>
            <span aria-hidden="true"> → </span>
            <time dateTime={report.request.endDate}>{report.request.endDate}</time>
            {selectedGroups && <> · {selectedGroups}</>}
          </p>
        </div>
        <dl className="report-meta">
          <div><dt>생성</dt><dd><time dateTime={report.generatedAtUtc}>{new Date(report.generatedAtUtc).toLocaleString("ko-KR", { timeZone: report.request.timeZoneId })}</time></dd></div>
          <div><dt>계산</dt><dd>{report.calculationVersion}</dd></div>
          <div><dt>Schema</dt><dd>{report.schemaVersion}</dd></div>
        </dl>
      </header>

      {report.status.state !== "complete" && (
        <aside className="root-status" role={report.status.state === "failed" ? "alert" : "status"}>
          <strong>{report.status.message ?? "리포트에 경고가 있습니다."}</strong>
          {[...report.status.warnings, ...report.status.errors].length > 0 && (
            <ul>{[...report.status.warnings, ...report.status.errors].map((issue, index) => <li key={`${issue.code}-${index}`}>{issue.message}</li>)}</ul>
          )}
        </aside>
      )}

      <ViewerToolbar preferences={preferences} defaults={report.viewerDefaults} onChange={setPreferences} />

      <div className="tab-strip">
        <div role="tablist" aria-label="BMES 리포트 탭" className="tab-list">
          {TAB_KEYS.map((tabKey, index) => {
            const selected = activeTab === tabKey;
            return (
              <button
                key={tabKey}
                ref={(element) => { tabRefs.current[index] = element; }}
                id={`tab-${tabKey}`}
                role="tab"
                type="button"
                aria-selected={selected}
                aria-controls={`panel-${tabKey}`}
                tabIndex={selected ? 0 : -1}
                className="tab-button"
                onClick={() => setActiveTab(tabKey)}
                onKeyDown={(event) => handleTabKey(event, index)}
              >
                {LABELS[tabKey]}
                {report.tabs[tabKey].status.state !== "complete" && <span className="status-dot" aria-label={report.tabs[tabKey].status.state} />}
              </button>
            );
          })}
        </div>
      </div>

      <section
        id={`panel-${activeTab}`}
        role="tabpanel"
        aria-labelledby={`tab-${activeTab}`}
        tabIndex={0}
        className="tab-panel"
      >
        <ReportTabContent report={report} tabKey={activeTab} preferences={preferences} />
      </section>
    </main>
  );
}

export { LABELS as TAB_LABELS };
