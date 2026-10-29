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
          : new ReportLoadError("network", "An unknown error occurred while loading the report.");
        setState({ kind: "error", error: reportError });
      });
    return () => controller.abort();
  }, [reportUrl, attempt]);

  if (state.kind === "loading") {
    return (
      <main className="center-state" aria-busy="true" aria-live="polite">
        <span className="loading-mark" aria-hidden="true" />
        <h1>Loading the BMES report</h1>
        <p>This may take a moment depending on the data size.</p>
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
    ? "The report link has expired"
    : error.kind === "unauthorized"
      ? "Sign-in required"
      : error.kind === "unsupported-schema"
        ? "Viewer update required"
        : "The report cannot be displayed";
  return (
    <main className="center-state error-card" role="alert">
      <p className="state-kicker">{error.kind}</p>
      <h1>{title}</h1>
      <p>{error.message}</p>
      <button type="button" className="primary-button" onClick={onRetry}>Retry</button>
    </main>
  );
}
