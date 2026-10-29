import { useEffect, useMemo, useState } from "react";
import { fetchDailyReport, refreshDailyReport } from "./api";
import type { DailyReportEnvelope } from "./contract";
import type { Granularity, GroupFilter, SortKey } from "./logic";
import { filterAndSortRows, selectGranularity } from "./logic";
import { Controls } from "./components/Controls";
import { DocHead } from "./components/DocHead";
import { GranularitySegment } from "./components/GranularitySegment";
import { ReadingsStrip } from "./components/ReadingsStrip";
import { RecordGrid } from "./components/RecordGrid";
import { RefreshState } from "./components/RefreshState";

export interface DailyReportViewProps {
  envelope: DailyReportEnvelope;
  onRefresh: () => void;
  notice?: string;
}

export function DailyReportView({ envelope, onRefresh, notice }: DailyReportViewProps) {
  const [granularity, setGranularity] = useState<Granularity>("week");
  const [group, setGroup] = useState<GroupFilter>("ALL");
  const [sort, setSort] = useState<SortKey>("delta");
  const snapshot = envelope.snapshot;

  const view = useMemo(() => {
    if (!snapshot) return null;
    const data = selectGranularity(snapshot, granularity);
    const filter = { group, query: "", sort };
    const worsened = filterAndSortRows(data.worsened, filter);
    const improved = filterAndSortRows(data.improved, filter);
    return { data, worsened, improved };
  }, [snapshot, granularity, group, sort]);

  if (!snapshot || snapshot.error) {
    return (
      <div className="dr-react-root">
        <RefreshState isRefreshing={envelope.isRefreshing} logs={envelope.logs} error={snapshot?.error || notice || null} onRefresh={onRefresh} />
      </div>
    );
  }

  return (
    <div className="dr-react-root">
      <main className="dr-sheet">
        <DocHead snapshot={snapshot} isRefreshing={envelope.isRefreshing} onRefresh={onRefresh} />
        {notice ? <div className="dr-api-notice" role="alert">{notice}</div> : null}
        <ReadingsStrip snapshot={snapshot} />
        <div className="dr-view-controls">
          <GranularitySegment value={granularity} onChange={setGranularity} />
          <Controls group={group} onGroupChange={setGroup} />
        </div>
        {view ? <RecordGrid {...view} onSortChange={setSort} /> : null}
        {envelope.isRefreshing && envelope.logs.length > 0 ? (
          <div className="dr-inline-log" aria-live="polite">{envelope.logs.slice(-1)[0]}</div>
        ) : null}
      </main>
    </div>
  );
}

export interface AppProps {
  initialEnvelope: DailyReportEnvelope;
  dataUrl: string;
  refreshUrl: string;
}

export function App({ initialEnvelope, dataUrl, refreshUrl }: AppProps) {
  const [envelope, setEnvelope] = useState(initialEnvelope);
  const [notice, setNotice] = useState("");

  useEffect(() => {
    if (!envelope.isRefreshing) return;
    let active = true;
    let timer = 0;
    const controller = new AbortController();

    const poll = () => {
      timer = window.setTimeout(() => {
        void fetchDailyReport(dataUrl, controller.signal)
          .then((next) => {
            if (!active) return;
            setEnvelope(next);
            setNotice("");
            if (next.isRefreshing) poll();
          })
          .catch((error: unknown) => {
            if (!active || (error instanceof DOMException && error.name === "AbortError")) return;
            setNotice(error instanceof Error ? error.message : "Could not update refresh progress.");
          });
      }, 2000);
    };
    poll();

    return () => {
      active = false;
      window.clearTimeout(timer);
      controller.abort();
    };
  }, [dataUrl, envelope.isRefreshing]);

  const refreshNow = () => {
    setNotice("");
    void refreshDailyReport(refreshUrl)
      .then(setEnvelope)
      .catch((error: unknown) => setNotice(error instanceof Error ? error.message : "Could not start refresh."));
  };
  return <DailyReportView envelope={envelope} onRefresh={refreshNow} notice={notice} />;
}
