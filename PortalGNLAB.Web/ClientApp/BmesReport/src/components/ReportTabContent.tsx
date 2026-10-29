import type { BmesReportDocument, FCostFollowerTabData, ReportStatus, TabKey } from "../contract";
import type { ViewerPreferences } from "../logic";
import { resolveFCostDataset } from "../logic";
import { CauseMonthlyTab } from "../tabs/CauseMonthlyTab";
import { DailyTab } from "../tabs/DailyTab";
import { FCostTab } from "../tabs/FCostTab";
import { KpiTab } from "../tabs/KpiTab";
import { WeeklyTab } from "../tabs/WeeklyTab";
import { EmptySection } from "./ReportTable";

export function ReportTabContent({
  report,
  tabKey,
  preferences,
}: {
  report: BmesReportDocument;
  tabKey: TabKey;
  preferences: ViewerPreferences;
}) {
  const status = report.tabs[tabKey].status;
  if (status.state === "failed") {
    return <TabFailure status={status} />;
  }

  let content: React.ReactNode;
  switch (tabKey) {
    case "daily": {
      const data = report.tabs.daily.data;
      content = data ? <DailyTab data={data} preferences={preferences} /> : <EmptySection />;
      break;
    }
    case "weekly": {
      const data = report.tabs.weekly.data;
      content = data ? <WeeklyTab data={data} preferences={preferences} /> : <EmptySection />;
      break;
    }
    case "kpi": {
      const data = report.tabs.kpi.data;
      content = data ? <KpiTab data={data} preferences={preferences} /> : <EmptySection />;
      break;
    }
    case "cause-monthly": {
      const data = report.tabs["cause-monthly"].data;
      content = data ? <CauseMonthlyTab data={data} preferences={preferences} /> : <EmptySection />;
      break;
    }
    case "fcost": {
      const data = report.tabs.fcost.data;
      content = data ? <FCostTab dataset={data.dataset} view={data.view} preferences={preferences} /> : <EmptySection />;
      break;
    }
    default: {
      const follower = report.tabs[tabKey].data as FCostFollowerTabData | null;
      const dataset = follower ? resolveFCostDataset(report, follower) : null;
      content = follower && dataset
        ? <FCostTab dataset={dataset} view={follower.view} preferences={preferences} />
        : <TabFailure status={{ ...status, message: "F-COST 원본 dataset을 찾을 수 없습니다." }} />;
      break;
    }
  }

  return (
    <>
      {status.state === "partial" && <IssueBanner status={status} />}
      {content}
    </>
  );
}

function TabFailure({ status }: { status: ReportStatus }) {
  const messages = status.errors.length > 0 ? status.errors.map((issue) => issue.message) : [status.message ?? "탭 데이터를 생성하지 못했습니다."];
  return (
    <div className="state-card error-card" role="alert">
      <h2>탭을 표시할 수 없습니다</h2>
      <ul>{messages.map((message, index) => <li key={`${index}-${message}`}>{message}</li>)}</ul>
      <p className="state-code">Code: {status.code}</p>
    </div>
  );
}

function IssueBanner({ status }: { status: ReportStatus }) {
  return (
    <aside className="issue-banner" aria-label="부분 데이터 경고">
      <strong>{status.message ?? "일부 데이터를 사용할 수 없습니다."}</strong>
      {status.warnings.length > 0 && (
        <ul>{status.warnings.map((issue, index) => <li key={`${issue.code}-${index}`}>{issue.message}</li>)}</ul>
      )}
    </aside>
  );
}
