import { useEffect, useState } from "react";
import { fetchReport, ReportLoadError } from "./api";
import type { BmesReportDocument } from "./contract";
import { ReportShell } from "./components/ReportShell";

type LoadState =
  | { kind: "loading" }
  | { kind: "ready"; report: BmesReportDocument }
  | { kind: "error"; error: ReportLoadError };

export function App({ reportUrl }: { reportUrl: string }) {
  const [attempt, setAttempt] = useState(0);
  const [state, setState] = useState<LoadState>({ kind: "loading" });

  useEffect(() => {
    const controller = new AbortController();
    setState({ kind: "loading" });
    fetchReport(reportUrl, controller.signal)
      .then((report) => setState({ kind: "ready", report }))
      .catch((error: unknown) => {
        if (error instanceof DOMException && error.name === "AbortError") return;
        const reportError = error instanceof ReportLoadError
          ? error
          : new ReportLoadError("network", "리포트를 불러오는 중 알 수 없는 오류가 발생했습니다.");
        setState({ kind: "error", error: reportError });
      });
    return () => controller.abort();
  }, [reportUrl, attempt]);

  if (state.kind === "loading") {
    return (
      <main className="center-state" aria-busy="true" aria-live="polite">
        <span className="loading-mark" aria-hidden="true" />
        <h1>BMES 리포트를 불러오는 중입니다</h1>
        <p>데이터 크기에 따라 잠시 걸릴 수 있습니다.</p>
      </main>
    );
  }

  if (state.kind === "error") {
    return <LoadError error={state.error} onRetry={() => setAttempt((current) => current + 1)} />;
  }

  return <ReportShell report={state.report} />;
}

function LoadError({ error, onRetry }: { error: ReportLoadError; onRetry: () => void }) {
  const title = error.kind === "expired"
    ? "리포트 링크가 만료되었습니다"
    : error.kind === "unauthorized"
      ? "로그인이 필요합니다"
      : error.kind === "unsupported-schema"
        ? "뷰어 업데이트가 필요합니다"
        : "리포트를 표시할 수 없습니다";
  return (
    <main className="center-state error-card" role="alert">
      <p className="state-kicker">{error.kind}</p>
      <h1>{title}</h1>
      <p>{error.message}</p>
      <button type="button" className="primary-button" onClick={onRetry}>다시 시도</button>
    </main>
  );
}
