import { useEffect, useRef, useState } from "react";
import { bmesReportApi, type ReportSetup, type Selection, type Stage, type StreamEvent } from "../bmesReportApi";
import "../bmesReport.css";
import { isoDate } from "./bmesReport/format";
import {
  entryLevelLabel, formatElapsed, logClass, pct, selectedCount, selectedInProductGroup, stageIcon, toggleGroup,
  toggleSubEntry, usDate, validateReport,
} from "./bmesReport/reportLogic";

const today = () => isoDate(new Date());
const plural = (n: number, word: string) => `${n} ${word}${n === 1 ? "" : "s"}`;

function StageRow({ stage }: { stage: Stage }) {
  const running = stage.status === "Running";
  return (
    <div className="br-stage">
      <div className="br-stage-head">
        <span>
          <span className={"br-stage-icon " + stage.status.toLowerCase()}>{stageIcon(stage)}</span> {stage.name}
        </span>
        <span className="br-stage-meta">
          {running && stage.stepTotal !== null && <span>{stage.phaseName} {stage.stepCurrent}/{stage.stepTotal}</span>}
          {running && stage.percent !== null && <span className="br-stage-pct">{stage.percent.toFixed(0)}%</span>}
          {stage.elapsedSeconds > 0 && <span>{formatElapsed(stage.elapsedSeconds)}</span>}
        </span>
      </div>
      <div className="br-bar thin">
        {running && stage.percent !== null && <div className="br-bar-fill" style={{ width: pct(stage.percent) }} />}
        {running && stage.percent === null && <div className="br-bar-fill busy" />}
        {stage.status === "Done" && <div className="br-bar-fill ok" />}
        {stage.status === "Failed" && <div className="br-bar-fill bad" />}
      </div>
      {running && stage.currentStep && <div className="br-stage-step" title={stage.currentStep}>{stage.currentStep}</div>}
    </div>
  );
}

/**
 * Report (port of BmesReportPage.razor): period + sub-group picker, per-stage progress and log while
 * BmesReportHtmlExportService builds every tab, then the generated report in the same iframe viewer
 * (/report/bmes/view/{token}, legacy HTML or the BmesReport React viewer).
 */
export default function BmesReport() {
  const [setup, setSetup] = useState<ReportSetup | null>(null);
  const [selection, setSelection] = useState<Selection>({ groupIds: [], subEntries: [] });
  const [activePg, setActivePg] = useState("ETC");
  const [start, setStart] = useState(today());
  const [end, setEnd] = useState(today());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [logs, setLogs] = useState<string[]>([]);
  const [stages, setStages] = useState<Stage[]>([]);
  const [overall, setOverall] = useState(0);
  const [success, setSuccess] = useState(false);
  const [collapsed, setCollapsed] = useState(false);
  const [logOpen, setLogOpen] = useState(false);
  const [token, setToken] = useState<string | null>(null);
  const [revision, setRevision] = useState(0);
  const logEnd = useRef<HTMLDivElement>(null);

  const applySetup = (s: ReportSetup) => {
    setSetup(s);
    setSelection(s.selection);
    setActivePg(s.activeProductGroup);
    setStages((current) => (current.length > 0 ? current : s.stageNames.map((name) => ({
      name, status: "Pending", currentStep: "", percent: null, stepCurrent: null, stepTotal: null, phaseName: null, elapsedSeconds: 0,
    }))));
  };

  useEffect(() => {
    bmesReportApi.report.setup().then(applySetup).catch((e: Error) => setError(e.message));
  }, []);

  useEffect(() => {
    logEnd.current?.scrollIntoView({ block: "nearest" });
  }, [logs.length, logOpen, collapsed]);

  const save = (next: Selection) => {
    setSelection(next);
    bmesReportApi.report.saveSelection(next).catch((e: Error) => setError(e.message));
  };

  async function refreshGroups() {
    try {
      await bmesReportApi.report.saveSelection(selection);
      applySetup(await bmesReportApi.report.setup());
    } catch (e) {
      setError((e as Error).message);
    }
  }

  async function getReport() {
    if (busy || !setup) return;
    const invalid = validateReport(start, end, setup.credentialsConfigured, selection);
    if (invalid) {
      setError(invalid);
      return;
    }
    setBusy(true);
    setSuccess(false);
    setError("");
    setLogs([]);
    setStages((s) => s.map((x) => ({ ...x, status: "Pending", currentStep: "", percent: null, stepCurrent: null, stepTotal: null, phaseName: null, elapsedSeconds: 0 })));
    setOverall(0);
    const onEvent = (event: StreamEvent) => {
      if (event.type === "log" || event.type === "progress") setLogs((l) => [...l, event.message]);
      else {
        setStages(event.stages);
        setOverall(event.overall);
      }
    };
    try {
      const result = await bmesReportApi.report.generate(start, end, selection, onEvent);
      setToken(result.token);
      setRevision((r) => r + 1);
      setSuccess(true);
    } catch (e) {
      setError("Error: " + (e as Error).message);
    } finally {
      setBusy(false);
    }
  }

  const groups = setup?.groups ?? [];
  const count = selectedCount(selection);
  const credentials = setup?.credentialsConfigured ?? false;

  return (
    <div className="br-page">
      <div className="br-head">
        <span className="br-icon">RP</span>
        <h1>Report</h1>
      </div>

      {collapsed ? (
        <button type="button" className="br-collapsed" onClick={() => setCollapsed(false)} title="Show report setup and progress log">
          <strong>Report Setup / Progress Log</strong>
          <span className="muted small">{usDate(start)} - {usDate(end)}</span>
          <span className="muted small grow-left">{busy ? "Processing..." : `${logs.length} log(s)`}</span>
          <span className="br-link">Show</span>
        </button>
      ) : (
        <div className="br-flow">
          <section className="br-card">
            <div className="br-card-head">
              <span>Report Setup</span>
              <button type="button" className="br-linkbtn" onClick={() => setCollapsed(true)} title="Hide report setup and progress log">Hide</button>
            </div>
            <div className="br-setup">
              <div className="br-setup-main">
                <div>
                  <div className="br-label">Period</div>
                  <div className="br-row">
                    <input type="date" className="inp sm" value={start} onChange={(e) => e.target.value && setStart(e.target.value)} />
                    <span className="muted">-</span>
                    <input type="date" className="inp sm" value={end} onChange={(e) => e.target.value && setEnd(e.target.value)} />
                    <button type="button" className="btn sm" onClick={() => { setStart(today()); setEnd(today()); }}>Today</button>
                  </div>
                </div>
                <div className="br-row">
                  <span className="muted small">BMES:</span>
                  {credentials ? (
                    <>
                      <span className="br-badge ok">Credentials OK</span>
                      <span className="muted small mono">{setup?.loginId}</span>
                    </>
                  ) : (
                    <>
                      <span className="br-badge warn">Not configured</span>
                      <a href="/app/setting" className="small">Setting</a>
                    </>
                  )}
                </div>
                {error && (
                  <div className="br-alert err">
                    <span className="br-badge bad">Error</span> {error}
                  </div>
                )}
              </div>

              <div className="br-groups">
                <div className="br-groups-head">
                  <span>Sub Groups</span>
                  <span className="br-row">
                    {count > 0 && <span className="br-pill">{count} selected</span>}
                    <span className="muted small">{plural(groups.length, "group")}</span>
                    <button type="button" className="br-linkbtn" onClick={refreshGroups} title="Refresh">Refresh</button>
                  </span>
                </div>
                <div className="br-picker">
                  {groups.length === 0 ? (
                    <div className="muted small br-pad">
                      No groups found. Define groups on the <a href="/app/bmes/make-model-group">Model Group</a> page.
                    </div>
                  ) : (
                    <>
                      <div className="br-pg-tabs">
                        {(setup?.productGroups ?? []).map((pg) => {
                          const total = groups.filter((g) => g.productGroup === pg).length;
                          const sel = selectedInProductGroup(groups, selection, pg);
                          return (
                            <button key={pg} type="button" className={"br-pg-tab" + (activePg === pg ? " active" : "")}
                              title={`${pg}: ${sel} selected of ${total}`} onClick={() => setActivePg(pg)}>
                              {pg} {sel > 0 ? <span>({sel})</span> : <span className="faint">{total}</span>}
                            </button>
                          );
                        })}
                      </div>
                      <div className="br-picker-list">
                        {groups.filter((g) => g.productGroup === activePg).map((g) => (
                          <div key={g.id}>
                            <label className="br-pick" title={g.name}>
                              <input type="checkbox" checked={selection.groupIds.includes(g.id)} onChange={() => save(toggleGroup(selection, g.id))} />
                              <span className="br-level">GROUP</span>
                              <span className="grow">{g.name}</span>
                              <span className="muted br-count">{g.lsCount} shifts</span>
                            </label>
                            {g.entries.map((entry) => (
                              <label key={entry.key} className="br-pick sub" title={entry.label} style={{ paddingLeft: 8 + entry.depth * 14 }}>
                                <input type="checkbox" checked={selection.subEntries.includes(entry.key)}
                                  onChange={() => save(toggleSubEntry(selection, entry.key))} />
                                <span className="br-level">{entryLevelLabel(entry)}</span>
                                <span className="grow">{entry.label}</span>
                                <span className="muted br-count">{entry.lsCount} shifts</span>
                              </label>
                            ))}
                          </div>
                        ))}
                        {!groups.some((g) => g.productGroup === activePg) && <div className="muted small br-pad center">No groups in {activePg}</div>}
                      </div>
                    </>
                  )}
                  <div className="br-picker-foot">
                    <button type="button" className="btn sm primary br-wide" onClick={getReport} disabled={busy || !credentials}>
                      {busy ? <><span className="br-spin" /> Processing...</> : <>Get Report{count > 0 && <span className="faint small"> ({plural(count, "group")})</span>}</>}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </section>

          <section className="br-card">
            <div className="br-card-head">
              <span>Progress / Log</span>
              <span className="br-row">
                {logs.length > 0 && <button type="button" className="br-linkbtn" onClick={() => { setLogs([]); setError(""); }}>Clear</button>}
                <button type="button" className="br-linkbtn" onClick={() => setCollapsed(true)} title="Hide report setup and progress log">Hide</button>
              </span>
            </div>
            <div className="br-pad">
              <div className="br-bar">
                {busy && stages.length > 0 && <div className="br-bar-fill" style={{ width: pct(overall) }} />}
                {busy && stages.length === 0 && <div className="br-bar-fill busy" />}
                {!busy && logs.length > 0 && <div className={"br-bar-fill " + (success ? "ok" : "bad")} />}
              </div>
              <div className="br-stages">
                {stages.map((s) => <StageRow key={s.name} stage={s} />)}
              </div>
              {stages.length > 0 && (
                <button type="button" className="br-linkbtn" onClick={() => setLogOpen((o) => !o)}>
                  {logOpen ? "Hide detail log" : `Show detail log (${logs.length})`}
                </button>
              )}
              {(stages.length === 0 || logOpen) && (
                <div className="br-log mono">
                  {logs.length === 0 ? <span className="faint">No log entries yet.</span> : logs.map((line, i) => <div key={i} className={logClass(line)}>{line}</div>)}
                  <div ref={logEnd} />
                </div>
              )}
            </div>
          </section>
        </div>
      )}

      <div className="br-viewer">
        {token === null ? (
          <div className="muted small br-pad-lg">
            Pick a period and sub groups, then press <strong>Get Report</strong> to build the report.
          </div>
        ) : (
          <iframe key={revision} src={`/report/bmes/view/${token}?rev=${revision}`} title="BMES Report" referrerPolicy="no-referrer" className="br-frame" />
        )}
      </div>
    </div>
  );
}
