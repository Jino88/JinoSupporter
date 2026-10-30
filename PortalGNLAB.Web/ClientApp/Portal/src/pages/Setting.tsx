import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  settingApi,
  type AiProviders,
  type CodexView,
  type Credentials,
  type HistoryView,
  type MaterialView,
  type PathsView,
  type QrDb,
  type QrModelCode,
  type SettingView,
} from "../settingApi";
import { defaultHistoryFrom, modelOptions, workFolders } from "../settingLogic";
import "../setting.css";

/** Port of Components/Pages/BmesSettingPage.razor (/setting). */

type Msg = { text: string; error: boolean } | null;

const errText = (e: unknown) => (e instanceof Error ? e.message : String(e));

function Note({ msg }: { msg: Msg }) {
  if (!msg?.text) return null;
  return <span className={"set-note " + (msg.error ? "err" : "ok")}>{msg.text}</span>;
}

function SecretInput({ value, onChange, placeholder, mono }: { value: string; onChange: (v: string) => void; placeholder?: string; mono?: boolean }) {
  const [show, setShow] = useState(false);
  return (
    <span className="set-secret">
      <input
        className={"inp sm" + (mono ? " mono" : "")}
        type={show ? "text" : "password"}
        value={value}
        placeholder={placeholder}
        autoComplete="off"
        onChange={(e) => onChange(e.target.value)}
      />
      <button type="button" className="btn sm" onClick={() => setShow(!show)}>
        {show ? "Hide" : "Show"}
      </button>
    </span>
  );
}

export default function Setting() {
  const [view, setView] = useState<SettingView>();
  const [loadError, setLoadError] = useState("");
  const [status, setStatus] = useState<Msg>(null);

  useEffect(() => {
    settingApi.load().then(setView).catch((e) => setLoadError(errText(e)));
  }, []);

  if (loadError) return <div className="alert">{loadError}</div>;
  if (!view) return <div className="muted">Loading...</div>;

  return (
    <div className="set-page">
      <header className="page-head">
        <h1>Setting</h1>
      </header>
      {status?.text && (
        <div className={"set-status " + (status.error ? "err" : "ok")}>
          <span>{status.text}</span>
          <button type="button" className="link" aria-label="Close" onClick={() => setStatus(null)}>
            ×
          </button>
        </div>
      )}

      {view.isAdmin && view.codex && view.aiProviders && view.qrDb && view.qrModels && (
        <>
          <div className="set-grid two">
            <CodexCard initial={view.codex} onProviders={(p) => setView({ ...view, aiProviders: p })} />
            <AiProvidersCard key={JSON.stringify(view.aiProviders)} initial={view.aiProviders} />
          </div>
          <QrDbCard initial={view.qrDb} />
          <QrModelsCard initial={view.qrModels.codes} path={view.qrModels.path} />
        </>
      )}

      <AppPathAndTables initial={view.paths} setStatus={setStatus} />
      <div className="set-grid two">
        <CredentialsCard initial={view.credentials} setStatus={setStatus} />
        <ApiLoginCard initial={view.credentials} setStatus={setStatus} />
      </div>
      <HistoryCard initial={view.history} setStatus={setStatus} />
      <MaterialCard initial={view.material} setStatus={setStatus} />
    </div>
  );
}

// ── Admin cards ─────────────────────────────────────────────────────────────────

function CodexCard({ initial, onProviders }: { initial: CodexView; onProviders: (p: AiProviders) => void }) {
  const [form, setForm] = useState(initial);
  const [models, setModels] = useState<string[]>([]);
  const [modelsLoading, setModelsLoading] = useState(false);
  const [modelsError, setModelsError] = useState("");
  const [showKey, setShowKey] = useState(false);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<Msg>(null);

  const loadModels = useCallback(async () => {
    setModelsLoading(true);
    setModelsError("");
    try {
      const r = await settingApi.codexModels();
      setModels(r.models);
      setModelsError(r.error);
    } catch (e) {
      setModelsError(`Model list unavailable: ${errText(e)}`);
    } finally {
      setModelsLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadModels();
  }, [loadModels]);

  async function save() {
    setSaving(true);
    setMsg(null);
    try {
      const r = await settingApi.saveCodex(form);
      setForm(r.codex);
      onProviders(r.aiProviders);
      setMsg({ text: "Saved.", error: false });
    } catch (e) {
      setMsg({ text: errText(e), error: true });
    } finally {
      setSaving(false);
    }
  }

  const set = (patch: Partial<CodexView>) => setForm({ ...form, ...patch });

  return (
    <section className="card">
      <h2>Codex API (OpenAI)</h2>
      <label className="set-field">
        API Key
        <span className="set-secret">
          <input
            className="inp sm mono"
            type={showKey ? "text" : "password"}
            placeholder="sk-..."
            value={form.apiKey}
            onChange={(e) => set({ apiKey: e.target.value })}
          />
          <button type="button" className="btn sm" onClick={() => setShowKey(!showKey)}>
            {showKey ? "Hide" : "Show"}
          </button>
        </span>
      </label>
      <label className="set-field">
        Translate Model
        <span className="set-secret">
          <select className="inp sm mono" value={form.translateModel} onChange={(e) => set({ translateModel: e.target.value })}>
            {modelOptions(models, form.translateModel).map((m) => (
              <option key={m} value={m}>
                {m}
              </option>
            ))}
          </select>
          <button type="button" className="btn sm" disabled={modelsLoading} title="Reload the model list from the OpenAI API" onClick={() => void loadModels()}>
            {modelsLoading ? "Loading..." : "Reload"}
          </button>
        </span>
        <span className={"set-hint" + (modelsError && !modelsLoading ? " err" : "")}>
          {modelsLoading ? "Loading model list..." : modelsError || `${models.length} model(s) available to this API key.`}
        </span>
      </label>
      <label className="set-field">
        Translate Reasoning Effort
        <select className="inp sm" value={form.translateEffort} onChange={(e) => set({ translateEffort: e.target.value })}>
          {form.reasoningEfforts.map((x) => (
            <option key={x} value={x}>
              {x}
            </option>
          ))}
        </select>
        <span className="set-hint">Codex Luna ({form.lunaModel}) at low is the default.</span>
      </label>
      <label className="set-field">
        Write (Review Report) Model
        <select className="inp sm mono" value={form.reportModel} onChange={(e) => set({ reportModel: e.target.value })}>
          {modelOptions(models, form.reportModel).map((m) => (
            <option key={m} value={m}>
              {m}
            </option>
          ))}
        </select>
      </label>
      <label className="set-field">
        Write (Review Report) Reasoning Effort
        <select className="inp sm" value={form.reportEffort} onChange={(e) => set({ reportEffort: e.target.value })}>
          {form.reasoningEfforts.map((x) => (
            <option key={x} value={x}>
              {x}
            </option>
          ))}
        </select>
        <span className="set-hint">
          Used by Write, Rewrite and Apply change. {form.defaultReportModel} at {form.defaultReportEffort} is the default.
        </span>
      </label>
      <div className="set-actions">
        <button type="button" className="btn sm primary" disabled={saving} onClick={() => void save()}>
          {saving ? "Saving..." : "Save"}
        </button>
        <Note msg={msg} />
      </div>
    </section>
  );
}

function AiProvidersCard({ initial }: { initial: AiProviders }) {
  const [form, setForm] = useState(initial);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<Msg>(null);

  async function save() {
    setSaving(true);
    setMsg(null);
    try {
      setForm(await settingApi.saveAiProviders({ codexCliEnabled: form.codexCliEnabled, codexApiEnabled: form.codexApiEnabled }));
      setMsg({ text: "Saved.", error: false });
    } catch (e) {
      setMsg({ text: errText(e), error: true });
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="card">
      <h2>AI Provider Switches</h2>
      <label className="set-check">
        <input type="checkbox" checked={form.codexCliEnabled} onChange={(e) => setForm({ ...form, codexCliEnabled: e.target.checked })} />
        Codex CLI
      </label>
      <label className="set-check">
        <input type="checkbox" checked={form.codexApiEnabled} onChange={(e) => setForm({ ...form, codexApiEnabled: e.target.checked })} />
        Codex API
      </label>
      <div className="set-actions">
        <button type="button" className="btn sm primary" disabled={saving} onClick={() => void save()}>
          {saving ? "Saving..." : "Save"}
        </button>
        <Note msg={msg} />
      </div>
    </section>
  );
}

function QrDbCard({ initial }: { initial: QrDb }) {
  const [form, setForm] = useState(initial);
  const [busy, setBusy] = useState<"" | "save" | "test">("");
  const [msg, setMsg] = useState<Msg>(null);
  const set = (patch: Partial<QrDb>) => setForm({ ...form, ...patch });

  async function run(action: "save" | "test") {
    setBusy(action);
    setMsg(null);
    try {
      if (action === "save") {
        const r = await settingApi.saveQrDb(form);
        setForm(r.qrDb);
        setMsg({ text: r.message, error: false });
      } else {
        const r = await settingApi.testQrDb(form);
        setForm(r.qrDb);
        setMsg({ text: r.message, error: r.error });
      }
    } catch (e) {
      setMsg({ text: errText(e), error: true });
    } finally {
      setBusy("");
    }
  }

  return (
    <section className="card">
      <h2>QR BAKO DATA DB</h2>
      <div className="form-row">
        <label>
          Server
          <input className="inp sm" value={form.server} onChange={(e) => set({ server: e.target.value })} />
        </label>
        <label>
          Port
          <input className="inp sm num" type="number" min={1} max={65535} value={form.port} onChange={(e) => set({ port: Number(e.target.value) })} />
        </label>
        <label>
          Database
          <input className="inp sm" value={form.database} onChange={(e) => set({ database: e.target.value })} />
        </label>
        <label>
          User ID
          <input className="inp sm" autoComplete="off" value={form.userId} onChange={(e) => set({ userId: e.target.value })} />
        </label>
        <label>
          Password
          <SecretInput value={form.password} onChange={(v) => set({ password: v })} />
        </label>
        <label>
          Rows
          <input className="inp sm num" type="number" min={1} max={20000} value={form.defaultMaxRows} onChange={(e) => set({ defaultMaxRows: Number(e.target.value) })} />
        </label>
        <label>
          Timeout
          <input className="inp sm num" type="number" min={1} max={300} value={form.timeoutSeconds} onChange={(e) => set({ timeoutSeconds: Number(e.target.value) })} />
        </label>
        <label className="set-check">
          <input type="checkbox" checked={form.encrypt} onChange={(e) => set({ encrypt: e.target.checked })} />
          Encrypt
        </label>
        <label className="set-check">
          <input type="checkbox" checked={form.trustServerCertificate} onChange={(e) => set({ trustServerCertificate: e.target.checked })} />
          Trust server certificate
        </label>
      </div>
      <div className="set-actions">
        <button type="button" className="btn sm primary" disabled={busy !== ""} onClick={() => void run("save")}>
          {busy === "save" ? "Saving..." : "Save"}
        </button>
        <button type="button" className="btn sm" disabled={busy !== ""} onClick={() => void run("test")}>
          {busy === "test" ? "Testing..." : "Test Connection"}
        </button>
        <Note msg={msg} />
      </div>
    </section>
  );
}

function QrModelsCard({ initial, path }: { initial: QrModelCode[]; path: string }) {
  const [rows, setRows] = useState(initial);
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState<Msg>(null);

  const update = (index: number, patch: Partial<QrModelCode>) => setRows(rows.map((r, i) => (i === index ? { ...r, ...patch } : r)));

  async function save() {
    setSaving(true);
    setMsg(null);
    try {
      // The server drops blank codes and duplicates; show what it actually kept.
      const r = await settingApi.saveQrModels(rows);
      setRows(r.codes);
      setMsg({ text: `${r.codes.length.toLocaleString("en-US")} model(s) saved.`, error: false });
    } catch (e) {
      setMsg({ text: errText(e), error: true });
    } finally {
      setSaving(false);
    }
  }

  return (
    <section className="card">
      <h2>
        QR Model Codes{" "}
        <span className="muted set-sub">
          · characters 11-12 of the QR name the model — J69NAG1GD0<b>37</b>00009 is code 37
        </span>
      </h2>
      <p className="muted small">
        The QR pages filter by this table. It is the only model marker the three stations agree on: BKTD writes{" "}
        <code>3200.MSM-X316.L1.M1.V01.260922</code> in <code>Model</code> while ALTD and GSTD write <code>3200.MSM-X316B.L1.M1.JIG1</code> in{" "}
        <code>ST_Information</code>.
      </p>
      <table className="set-table">
        <thead>
          <tr>
            <th style={{ width: "7rem" }}>Code</th>
            <th>Model name</th>
            <th style={{ width: "6rem" }}></th>
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 && (
            <tr>
              <td colSpan={3} className="muted small">
                No model is defined; the QR pages will only offer "All models".
              </td>
            </tr>
          )}
          {rows.map((row, i) => (
            <tr key={i}>
              <td>
                <input className="inp sm mono" maxLength={2} placeholder="37" value={row.code} onChange={(e) => update(i, { code: e.target.value })} />
              </td>
              <td>
                <input className="inp sm" placeholder="MSM-X316B" value={row.name} onChange={(e) => update(i, { name: e.target.value })} />
              </td>
              <td>
                <button type="button" className="btn sm danger" onClick={() => setRows(rows.filter((_, j) => j !== i))}>
                  Remove
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <div className="set-actions">
        <button type="button" className="btn sm" onClick={() => setRows([...rows, { code: "", name: "" }])}>
          Add model
        </button>
        <button type="button" className="btn sm primary" disabled={saving} onClick={() => void save()}>
          {saving ? "Saving..." : "Save"}
        </button>
        <Note msg={msg} />
        <span className="muted small set-right">{path}</span>
      </div>
    </section>
  );
}

// ── Every Setting user ──────────────────────────────────────────────────────────

type SetStatus = (msg: Msg) => void;

/** App Path and Server Tables share the typed working folder: loading a table saves it first. */
function AppPathAndTables({ initial, setStatus }: { initial: PathsView; setStatus: SetStatus }) {
  const [paths, setPaths] = useState(initial);
  const [root, setRoot] = useState(initial.dbRootDirectory);
  const [pathBusy, setPathBusy] = useState(false);
  const [pathMsg, setPathMsg] = useState<Msg>(null);
  const [restart, setRestart] = useState(false);
  const [tableBusy, setTableBusy] = useState<"" | "routing" | "reason" | "all">("");
  const [tableMsg, setTableMsg] = useState<Msg>(null);
  const folders = useMemo(() => workFolders(root, paths), [root, paths]);

  async function savePath() {
    setPathBusy(true);
    setPathMsg(null);
    try {
      const r = await settingApi.saveAppPath(root);
      setPaths(r.paths);
      setRoot(r.paths.dbRootDirectory);
      setRestart(r.requiresRestart);
      setPathMsg({ text: r.message, error: false });
    } catch (e) {
      setPathMsg({ text: errText(e), error: true });
    } finally {
      setPathBusy(false);
    }
  }

  async function pull(table: "routing" | "reason" | "all") {
    setTableBusy(table);
    setTableMsg(null);
    try {
      const r = await settingApi.pullServerTable(table, root);
      setPaths(r.paths);
      setRoot(r.paths.dbRootDirectory);
      setRestart(r.requiresRestart);
      setTableMsg({ text: r.message, error: !r.succeeded });
      setStatus({ text: r.succeeded ? r.message : `Server table load failed: ${r.message}`, error: !r.succeeded });
    } catch (e) {
      setTableMsg({ text: errText(e), error: true });
      setStatus({ text: `Server table load failed: ${errText(e)}`, error: true });
    } finally {
      setTableBusy("");
    }
  }

  const busy = pathBusy || tableBusy !== "";

  return (
    <>
      <section className="card">
        <h2>App Path</h2>
        <label className="set-field">
          <span>
            Working Folder <span className="tag set-warn">restart</span>
          </span>
          <span className="set-secret">
            <input className="inp sm mono" value={root} placeholder={paths.defaultDbRootDirectory} onChange={(e) => setRoot(e.target.value)} />
            <button
              type="button"
              className="btn sm"
              onClick={() => {
                setRoot(paths.defaultDbRootDirectory);
                setPathMsg({ text: "Default working folder applied. Click Save to keep it.", error: false });
              }}
            >
              Default
            </button>
            <button type="button" className="btn sm primary" disabled={busy} onClick={() => void savePath()}>
              {pathBusy ? "Saving..." : "Save"}
            </button>
          </span>
        </label>
        <div className="muted small set-paths">
          NG RATE: <code>{folders.ngRate}</code>
          <br />
          FCOST: <code>{folders.fCost}</code>
          <br />
          WK STATUS: <code>{folders.workerStatus}</code>
        </div>
        {pathMsg?.text && <div className={"set-line " + (pathMsg.error ? "err" : "ok")}>{pathMsg.text}</div>}
        {restart && <div className="set-line warn">Restart the app for main DB path changes to fully take effect.</div>}
      </section>

      <section className="card">
        <h2>Server Tables</h2>
        <div className="muted small">
          Save target: <code>{folders.serverTableTarget}</code>
        </div>
        <div className="set-actions">
          <button type="button" className="btn sm" disabled={busy} onClick={() => void pull("routing")}>
            {tableBusy === "routing" ? "Loading..." : "Load Routing"}
          </button>
          <button type="button" className="btn sm" disabled={busy} onClick={() => void pull("reason")}>
            {tableBusy === "reason" ? "Loading..." : "Load Reason"}
          </button>
          <button type="button" className="btn sm primary" disabled={busy} onClick={() => void pull("all")}>
            {tableBusy === "all" ? "Loading..." : "Load Both"}
          </button>
        </div>
        {tableMsg?.text && <div className={"set-line " + (tableMsg.error ? "err" : "ok")}>{tableMsg.text}</div>}
      </section>
    </>
  );
}

function CredentialsCard({ initial, setStatus }: { initial: Credentials; setStatus: SetStatus }) {
  const [loginId, setLoginId] = useState(initial.loginId);
  const [password, setPassword] = useState(initial.password);

  async function save() {
    try {
      setStatus({ text: (await settingApi.saveCredentials({ loginId, password })).message, error: false });
    } catch (e) {
      setStatus({ text: errText(e), error: true });
    }
  }

  return (
    <section className="card">
      <h2>
        BMES Credentials <span className="muted set-sub">(NG Rate / Worker Status)</span>
      </h2>
      <label className="set-field">
        Login ID
        <input className="inp sm" placeholder="BMES ID" value={loginId} onChange={(e) => setLoginId(e.target.value)} />
      </label>
      <label className="set-field">
        Password
        <input className="inp sm" type="password" placeholder="BMES Password" value={password} onChange={(e) => setPassword(e.target.value)} />
      </label>
      <button type="button" className="btn sm primary" onClick={() => void save()}>
        Save
      </button>
    </section>
  );
}

function ApiLoginCard({ initial, setStatus }: { initial: Credentials; setStatus: SetStatus }) {
  const [apiFacility, setFacility] = useState(initial.apiFacility);
  const [apiLoginId, setLoginId] = useState(initial.apiLoginId);
  const [apiPassword, setPassword] = useState(initial.apiPassword);

  async function save() {
    try {
      setStatus({ text: (await settingApi.saveApiCredentials({ apiFacility, apiLoginId, apiPassword })).message, error: false });
    } catch (e) {
      setStatus({ text: errText(e), error: true });
    }
  }

  return (
    <section className="card">
      <h2>
        BMES API Login <span className="muted set-sub">(PLATFORM_API)</span>
      </h2>
      <label className="set-field">
        Facility
        <select className="inp sm" value={apiFacility} onChange={(e) => setFacility(e.target.value)}>
          <option value="">Select facility…</option>
          <option value="BN">BN</option>
          <option value="GN">GN</option>
        </select>
      </label>
      <label className="set-field">
        Login ID
        <input className="inp sm" placeholder="BMES API ID" value={apiLoginId} onChange={(e) => setLoginId(e.target.value)} />
      </label>
      <label className="set-field">
        Password
        <input className="inp sm" type="password" placeholder="BMES API Password" value={apiPassword} onChange={(e) => setPassword(e.target.value)} />
      </label>
      <button type="button" className="btn sm primary" onClick={() => void save()}>
        Save
      </button>
    </section>
  );
}

/** The run lives on the server and keeps going without this page; the card polls while it runs. */
function HistoryCard({ initial, setStatus }: { initial: HistoryView; setStatus: SetStatus }) {
  const [history, setHistory] = useState(initial);
  const [from, setFrom] = useState(defaultHistoryFrom());
  const logRef = useRef<HTMLPreElement>(null);

  useEffect(() => {
    if (!history.isRunning) return;
    const timer = window.setInterval(() => {
      settingApi.history().then(setHistory).catch(() => undefined);
    }, 2000);
    return () => window.clearInterval(timer);
  }, [history.isRunning]);

  useEffect(() => {
    if (logRef.current) logRef.current.scrollTop = logRef.current.scrollHeight;
  }, [history.log]);

  async function start() {
    try {
      setHistory(await settingApi.startHistory(from));
    } catch (e) {
      setStatus({ text: errText(e), error: true });
    }
  }

  async function cancel() {
    try {
      setHistory(await settingApi.cancelHistory());
    } catch (e) {
      setStatus({ text: errText(e), error: true });
    }
  }

  return (
    <section className="card">
      <h2>
        History Cache <span className="muted set-sub">(NG Rate / F-COST)</span>
        {history.isRunning ? <span className="tag">Running</span> : history.status ? <span className="tag set-grey">{history.status}</span> : null}
      </h2>
      <p className="muted small">
        Fetches every day from the start date to <strong>{history.lastHistoryDay}</strong> from the BMES API in one run and keeps it. The last{" "}
        {history.liveWindowDays} days are never kept: NG Rate and F-COST always fetch them live. Days already cached are skipped, so a stopped run
        continues where it left off. It keeps running if you leave this page.
      </p>
      <div className="form-row">
        <label>
          From
          <input className="inp sm" type="date" value={from} disabled={history.isRunning} onChange={(e) => setFrom(e.target.value)} />
        </label>
        {history.isRunning ? (
          <button type="button" className="btn sm danger" onClick={() => void cancel()}>
            Stop
          </button>
        ) : (
          <button type="button" className="btn sm primary" disabled={!from} onClick={() => void start()}>
            ▶ Build history cache
          </button>
        )}
      </div>
      {history.log.length > 0 && (
        <pre ref={logRef} className="set-log">
          {history.log.join("\n")}
        </pre>
      )}
    </section>
  );
}

function MaterialCard({ initial, setStatus }: { initial: MaterialView; setStatus: SetStatus }) {
  const [material, setMaterial] = useState(initial);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);

  async function fetchAll() {
    setBusy(true);
    setLog([]);
    try {
      const last = await settingApi.fetchMaterials((event) => {
        if (event.log !== undefined) setLog((lines) => [...lines, event.log!]);
      });
      if (last?.material) setMaterial(last.material);
      if (last?.error) setStatus({ text: last.error, error: true });
      else if ((last?.saved ?? 0) > 0)
        setStatus({ text: `Saved ${last!.saved} material(s). Total in DB: ${last!.material?.count ?? material.count}`, error: false });
      else setStatus({ text: "Fetch failed — see log below.", error: true });
    } catch (e) {
      setStatus({ text: `Unexpected error: ${errText(e)}`, error: true });
    } finally {
      setBusy(false);
    }
  }

  const m = material.latest;
  return (
    <section className="card">
      <h2 className="set-head">
        <span>🧪 BMES Material Fetch (MES020010)</span>
        <span className="tag set-grey">DB rows: {material.count}</span>
        <button type="button" className="btn sm primary set-right" disabled={busy} onClick={() => void fetchAll()}>
          {busy ? "Fetching…" : "▶ Fetch all materials"}
        </button>
      </h2>
      {log.length > 0 && <pre className="set-log">{log.join("\n")}</pre>}
      {m ? (
        <div className="small set-material">
          <div>
            <strong>MATNR:</strong> {m.matnr}
          </div>
          <div>
            <strong>MAKTX:</strong> {m.maktx}
          </div>
          <div>
            <strong>MTYPE/BTYPE:</strong> {m.mtype} / {m.btype}
          </div>
          <div>
            <strong>ModNameB:</strong> {m.modNameB}
          </div>
          <div>
            <strong>LotQt / Bunch:</strong> {m.lotQt} / {m.bunch}
          </div>
          <div>
            <strong>ERNAM / ERDAT:</strong> {m.ernam} / {m.erdat}
          </div>
          <div className="muted">FetchedAt: {m.fetchedAt}</div>
        </div>
      ) : (
        <div className="muted small">No material saved yet.</div>
      )}
    </section>
  );
}
