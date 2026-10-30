import { useCallback, useEffect, useState } from "react";
import { ngRateApi, RunError, type NgRateSetup } from "../../ngRateApi";
import { stampLog, today } from "./logic";

/**
 * State every NG Rate page shares: the setup (credentials + Model Groups), the period, the progress
 * log and the busy / error flags of one Get Report run.
 */
export function useReportRun() {
  const [setup, setSetup] = useState<NgRateSetup | null>(null);
  const [start, setStart] = useState(today());
  const [end, setEnd] = useState(today());
  const [busy, setBusy] = useState(false);
  const [logs, setLogs] = useState<string[]>([]);
  const [savedDbPath, setSavedDbPath] = useState<string | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [reportError, setReportError] = useState("");

  const loadSetup = useCallback(async (): Promise<NgRateSetup | null> => {
    try {
      const next = await ngRateApi.setup();
      setSetup(next);
      return next;
    } catch (e) {
      setErrorMessage((e as Error).message);
      return null;
    }
  }, []);

  useEffect(() => {
    void loadSetup();
  }, [loadSetup]);

  const setToday = useCallback(() => {
    setStart(today());
    setEnd(today());
  }, []);

  /** Checks the period and credentials; returns false (with the message shown) when the run cannot start. */
  const canRun = (): boolean => {
    if (end < start) {
      setReportError("End date must be on or after start date.");
      return false;
    }
    if (!setup?.credentialsConfigured) {
      setReportError("BMES credentials are not configured.");
      return false;
    }
    return true;
  };

  /** Runs one report call with a fresh log; null when it failed (the message is already shown). */
  async function run<T extends { dbPath: string }>(call: (onProgress: (message: string) => void) => Promise<T>): Promise<T | null> {
    setReportError("");
    setErrorMessage(null);
    setSavedDbPath(null);
    setLogs([]);
    setBusy(true);
    try {
      const result = await call((message) => setLogs((current) => [...current, stampLog(message)]));
      setSavedDbPath(result.dbPath);
      return result;
    } catch (e) {
      const message = (e as Error).message;
      if (e instanceof RunError && e.panel) setErrorMessage(message);
      setReportError(message);
      return null;
    } finally {
      setBusy(false);
    }
  }

  const clearResult = () => {
    setSavedDbPath(null);
    setErrorMessage(null);
  };

  return {
    setup, loadSetup, start, setStart, end, setEnd, setToday, busy, logs, clearLogs: () => setLogs([]),
    savedDbPath, errorMessage, clearResult, reportError, setReportError, canRun, run,
  };
}
