interface RefreshStateProps {
  isRefreshing: boolean;
  logs: string[];
  error?: string | null;
  onRefresh: () => void;
}

export function RefreshState({ isRefreshing, logs, error, onRefresh }: RefreshStateProps) {
  return (
    <div className={error ? "dr-state dr-state-error" : "dr-state"}>
      {isRefreshing ? <span className="dr-spinner dr-spinner-large" /> : null}
      <strong>{error ? "Could not build the report." : isRefreshing ? "Aggregating the last two months from BMES." : "No report is cached yet."}</strong>
      {error ? <span>{error}</span> : <span>{isRefreshing ? "This can take a few minutes." : "Build the first snapshot to continue."}</span>}
      {!isRefreshing ? <button className="dr-action-button" type="button" onClick={onRefresh}>Refresh now</button> : null}
      {logs.length > 0 ? (
        <div className="dr-refresh-log" aria-live="polite">
          {logs.slice(-6).map((line, index) => <div key={`${line}-${index}`}>{line}</div>)}
        </div>
      ) : null}
    </div>
  );
}
