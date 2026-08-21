import type { DailyReportSnapshot } from "../contract";

interface DocHeadProps {
  snapshot: DailyReportSnapshot;
  isRefreshing: boolean;
  onRefresh: () => void;
}

export function DocHead({ snapshot, isRefreshing, onRefresh }: DocHeadProps) {
  const generated = snapshot.generatedAt.slice(5, 16).replace("T", " ");
  return (
    <header className="dr-doc-head">
      <div className="dr-doc-copy">
        <div className="dr-doc-kicker">BMES · DEFECT RATE TREND BY MODEL</div>
        <h1 className="dr-doc-title">DAILY REPORT</h1>
        <div className="dr-doc-meta">
          <span>{snapshot.startDate.slice(0, 10)} – {snapshot.endDate.slice(0, 10)}</span>
          <span className="dr-doc-sep">·</span>
          <span>AS OF {generated}</span>
          <span className="dr-doc-sep">·</span>
          <span>{snapshot.modelCount} MODELS</span>
        </div>
      </div>
      <div className="dr-doc-actions">
        {isRefreshing ? <span className="dr-refresh-label"><span className="dr-spinner" />Fetching from BMES…</span> : null}
        <button className="dr-action-button" type="button" onClick={onRefresh} disabled={isRefreshing}>
          Refresh now
        </button>
        <a className="dr-action-button dr-action-link" href="/report/bmes">Full report →</a>
      </div>
    </header>
  );
}
