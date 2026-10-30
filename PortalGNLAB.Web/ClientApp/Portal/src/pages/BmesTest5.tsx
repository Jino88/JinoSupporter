import { useEffect, useMemo, useState } from "react";
import {
  bmesReportApi, type BomResult, type CachedModel, type CostRow, type ModelCatalog, type PdmRow,
} from "../bmesReportApi";
import { downloadBlob } from "../toolsApi";
import "../bmesReport.css";
import { base64ToBlob, isoDate, n0, stamp } from "./bmesReport/format";
import {
  buildStandaloneHtml, byProduct, displayName, filterCachedModels, findPdm, formatUnitPrice, formatVnd, localSuggestions,
  matchedProducts, pdmIndex, sanitize, visibleRows, type Suggestion,
} from "./bmesReport/test5Logic";

const emptyCatalog: ModelCatalog = { models: [], syncedAt: "", workDate: "" };

/**
 * BOM & Drawing (port of BmesTest5Page.razor): pick cached models or search the local BMES model-name
 * catalog, load each BOM from the local cache or the network DB (SELECT only), show PDM drawing numbers
 * and unit prices inline, download a DWG or the whole BOM as HTML.
 */
export default function BmesTest5() {
  const [catalog, setCatalog] = useState<ModelCatalog>(emptyCatalog);
  const [cached, setCached] = useState<CachedModel[]>([]);
  const [pdmTotal, setPdmTotal] = useState(0);
  const [pdmFetchedAt, setPdmFetchedAt] = useState("");

  const [filter, setFilter] = useState("");
  const [suggestions, setSuggestions] = useState<Suggestion[]>([]);
  const [searchHasRun, setSearchHasRun] = useState(false);
  const [suppress, setSuppress] = useState(false);
  const [syncing, setSyncing] = useState(false);
  const [searchError, setSearchError] = useState("");

  const [maxDepth, setMaxDepth] = useState(6);
  const [plant, setPlant] = useState("3200");
  const [workDate, setWorkDate] = useState(isoDate(new Date()));
  const [hideNoise, setHideNoise] = useState(true);
  const [fromServer, setFromServer] = useState(false);
  const [selected, setSelected] = useState<Set<string>>(new Set());

  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");
  const [error, setError] = useState("");
  const [results, setResults] = useState<BomResult[]>([]);
  const [usedOptions, setUsedOptions] = useState({ plant: "3200", maxDepth: 6, workDate: "" });

  const [pdm, setPdm] = useState<Map<string, PdmRow>>(new Map());
  const [pdmBusy, setPdmBusy] = useState(false);
  const [downloadBusy, setDownloadBusy] = useState(false);
  const [pdmStatus, setPdmStatus] = useState("");
  const [pdmError, setPdmError] = useState(false);

  const [costs, setCosts] = useState<Map<string, CostRow>>(new Map());
  const [costBusy, setCostBusy] = useState(false);
  const [costError, setCostError] = useState("");

  const cachedNames = useMemo(() => new Set(cached.map((m) => m.name.toUpperCase())), [cached]);
  const shownModels = useMemo(() => filterCachedModels(cached, filter), [cached, filter]);
  const selectedList = useMemo(() => [...selected].sort((a, b) => a.toUpperCase() < b.toUpperCase() ? -1 : 1), [selected]);
  const visible = useMemo(() => results.map((r) => visibleRows(r.rows, hideNoise)), [results, hideNoise]);

  useEffect(() => {
    bmesReportApi.test5.state(plant).then((s) => {
      setCached(s.cachedModels);
      setCatalog(s.catalog);
      setPdmTotal(s.pdmTotal);
      setPdmFetchedAt(s.pdmFetchedAt);
    }).catch((e: Error) => setError("Model list load failed: " + e.message));
  }, []);

  /** Plant is committed on change (blur / Enter), like the Blazor @bind. */
  async function changePlant(next: string) {
    try {
      const c = await bmesReportApi.test5.catalog(next);
      setCatalog(c);
      setSuggestions(localSuggestions(filter.trim(), c.models, cached));
    } catch (e) {
      setSearchError("Local model-name cache load failed: " + (e as Error).message);
    }
  }

  function onSearch(value: string) {
    setFilter(value);
    setSuppress(false);
    setSearchError("");
    setSearchHasRun(value.trim().length >= 2);
    setSuggestions(localSuggestions(value.trim(), catalog.models, cached));
  }

  async function syncCatalog() {
    if (busy || syncing) return;
    setSyncing(true);
    setSearchError("");
    setSuppress(false);
    try {
      const c = await bmesReportApi.test5.syncCatalog(plant, workDate);
      setCatalog(c);
      setSuggestions(localSuggestions(filter.trim(), c.models, cached));
      setSearchHasRun(filter.trim().length >= 2);
      setStatus(`${n0(c.models.length)} BMES model names saved locally. Typing now uses the local cache only.`);
    } catch (e) {
      setSearchError("BMES model-name synchronization failed: " + (e as Error).message);
      setSearchHasRun(true);
    } finally {
      setSyncing(false);
    }
  }

  async function refreshPdmMatches(list: BomResult[], rowsByResult: ReturnType<typeof visibleRows>[], total: number) {
    if (total === 0 || list.length === 0) {
      setPdm(new Map());
      return 0;
    }
    try {
      const names = rowsByResult.flat().map(displayName);
      const { rows } = await bmesReportApi.test5.pdmMatch(names, list.map((r) => r.model));
      setPdm(pdmIndex(rows));
      return rows.length;
    } catch (e) {
      setPdm(new Map());
      setPdmError(true);
      setPdmStatus("PDM match failed: " + (e as Error).message);
      return 0;
    }
  }

  async function loadCosts(rowsByResult: ReturnType<typeof visibleRows>[]) {
    setCosts(new Map());
    setCostError("");
    const codes = [...new Set(rowsByResult.flat().map((r) => r.materialCode.trim()).filter((c) => c))];
    if (codes.length === 0) return;
    setCostBusy(true);
    try {
      const answer = await bmesReportApi.test5.costs(codes);
      setCosts(new Map(answer.rows.map((r) => [r.materialCode.toUpperCase(), r])));
      setCostError(answer.error);
    } catch (e) {
      setCostError("Unit price lookup failed: " + (e as Error).message);
    } finally {
      setCostBusy(false);
    }
  }

  async function loadBom(models: string[]) {
    if (models.length === 0) return;
    setBusy(true);
    setError("");
    setResults([]);
    setStatus("");
    try {
      const answer = await bmesReportApi.test5.loadBom({ models, plant, workDate, maxDepth, fromServer }, setStatus);
      const rowsByResult = answer.results.map((r) => visibleRows(r.rows, hideNoise));
      const total = rowsByResult.reduce((s, r) => s + r.length, 0);
      const failed = answer.results.filter((r) => r.error.trim()).length;
      setResults(answer.results);
      setUsedOptions({ plant, maxDepth, workDate });
      setCached(answer.cachedModels);
      setStatus(failed === 0
        ? `${answer.results.length} model(s), ${n0(total)} BOM rows — cache ${answer.cacheLoads}, server ${answer.serverLoads}.`
        : `${answer.results.length} model(s), ${n0(total)} BOM rows, ${failed} failed — cache ${answer.cacheLoads}, server ${answer.serverLoads}.`);
      await refreshPdmMatches(answer.results, rowsByResult, pdmTotal);
      await loadCosts(rowsByResult);
    } catch (e) {
      setError("BOM load failed: " + (e as Error).message);
      setStatus("");
    } finally {
      setBusy(false);
    }
  }

  function loadSuggested(s: Suggestion) {
    if (busy || !s.name.trim()) return;
    setFilter(s.name);
    setSearchHasRun(false);
    setSuggestions([]);
    setSuppress(true);
    setSelected(new Set([s.name]));
    loadBom([s.name]);
  }

  function toggleModel(name: string, on: boolean) {
    // The Blazor page kept a case-insensitive set.
    setSelected((cur) => {
      const next = new Set([...cur].filter((s) => s.toUpperCase() !== name.toUpperCase()));
      if (on) next.add(name);
      return next;
    });
  }

  async function fetchPdm() {
    if (pdmBusy) return;
    setPdmBusy(true);
    setPdmError(false);
    setPdmStatus("");
    try {
      const r = await bmesReportApi.test5.fetchPdm(setPdmStatus);
      if (r.saved < 0) {
        setPdmError(true);
      } else {
        setPdmTotal(r.total);
        setPdmFetchedAt(r.fetchedAt);
        await refreshPdmMatches(results, visible, r.total);
        setPdmStatus(`✓ ${n0(r.saved)} PDM row(s) saved.`);
      }
    } catch (e) {
      setPdmError(true);
      setPdmStatus("✗ " + (e as Error).message);
    } finally {
      setPdmBusy(false);
    }
  }

  async function downloadDrawing(pdmno: string, pdmvr: string) {
    if (downloadBusy) return;
    setDownloadBusy(true);
    setPdmError(false);
    setPdmStatus(`Downloading drawing ${pdmno}…`);
    try {
      const r = await bmesReportApi.test5.drawing(pdmno, pdmvr, setPdmStatus);
      if (!r.success) {
        setPdmError(true);
        setPdmStatus(`✗ ${pdmno}: ${r.error}`);
        return;
      }
      const blob = base64ToBlob(r.base64, "application/octet-stream");
      downloadBlob(blob, r.fileName);
      setPdmStatus(`✓ ${r.fileName} (${n0(blob.size / 1024)} KB)`);
    } catch (e) {
      setPdmError(true);
      setPdmStatus(`✗ ${pdmno}: ${(e as Error).message}`);
    } finally {
      setDownloadBusy(false);
    }
  }

  function downloadHtml() {
    if (results.length === 0) return;
    const now = new Date();
    const html = buildStandaloneHtml(results, pdm, {
      ...usedOptions,
      hideNoise,
      generated: `${isoDate(now)} ${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`,
    });
    const label = results.length === 1 ? results[0].model : `${results.length}-models`;
    const name = `bom-${sanitize(label)}-${stamp(now, false, "-")}.html`;
    downloadBlob(new Blob([html], { type: "text/html; charset=utf-8" }), name);
    bmesReportApi.test5.activity("Download HTML " + name).catch(() => undefined);
  }

  const catalogTitle = catalog.models.length === 0
    ? "Fetch BOM-enabled P-S, P-M, P-N, and P-H product codes and model names from BMES, then search them locally."
    : `${n0(catalog.models.length)} local model names · synced ${catalog.syncedAt} · BOM date ${catalog.workDate}`;
  const showSuggestions = !suppress && (syncing || suggestions.length > 0 || searchHasRun);

  return (
    <div className="br-page t5">
      <div className="br-head">
        <span className="br-icon">B&amp;D</span>
        <h1>BOM &amp; Drawing</h1>
      </div>

      {error && <div className="br-alert err">{error}</div>}

      <div className="t5-toolbar">
        <div className="t5-group">
          <span className="muted small">Model {catalog.models.length > 0 && <span>({n0(catalog.models.length)} local names)</span>}</span>
          <input className="inp sm" placeholder="Type a similar model name" autoComplete="off" value={filter}
            onChange={(e) => onSearch(e.target.value)} disabled={busy || syncing} />
          {showSuggestions && (
            <span className="t5-suggestions">
              {syncing && <span className="t5-suggestion-status"><span className="br-spin" /> Synchronizing model names from BMES once...</span>}
              {suggestions.map((s) => (
                <button key={s.source + s.name} type="button" className="t5-suggestion" title={s.hint} onClick={() => loadSuggested(s)}>
                  <span className="t5-suggestion-main">
                    <span>{s.name}</span>
                    {s.productCode && <small>{s.productCode}</small>}
                  </span>
                  {cachedNames.has(s.name.toUpperCase())
                    ? <span className="br-badge ok">cached</span>
                    : s.source === "catalog" ? <span className="br-badge info">local</span> : null}
                </button>
              ))}
              {!syncing && suggestions.length === 0 && (
                <span className={"t5-suggestion-status" + (searchError ? " err" : "")}>
                  {searchError || (catalog.models.length === 0 ? "No local model-name cache. Click Sync model names." : "No cached model name matches.")}
                </span>
              )}
              {!syncing && suggestions.length > 0 && searchError && <span className="t5-suggestion-status err">{searchError}</span>}
            </span>
          )}
        </div>

        <button type="button" className="btn sm" onClick={syncCatalog} disabled={busy || syncing} title={catalogTitle}>
          {syncing ? "Syncing names…" : "Sync model names"}
        </button>
        <label className="t5-field narrow">
          <span className="muted small">Depth</span>
          <input type="number" className="inp sm" min={1} max={12} value={maxDepth}
            onChange={(e) => setMaxDepth(Number(e.target.value) || 1)} disabled={busy} />
        </label>
        <label className="t5-field narrow">
          <span className="muted small">Plant</span>
          <input className="inp sm" value={plant} onChange={(e) => setPlant(e.target.value)} onBlur={(e) => changePlant(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && changePlant(plant)} disabled={busy || syncing} />
        </label>
        <label className="t5-field">
          <span className="muted small">Work date</span>
          <input type="date" className="inp sm" value={workDate} onChange={(e) => e.target.value && setWorkDate(e.target.value)} disabled={busy} />
        </label>
        <label className="t5-check">
          <input type="checkbox" checked={hideNoise} onChange={(e) => setHideNoise(e.target.checked)} disabled={busy} />
          <span className="small">Hide M-P / C-S codes</span>
        </label>
        <label className="t5-check" title="Unchecked: use the local model cache first. Cache misses are fetched from BMES automatically.">
          <input type="checkbox" checked={fromServer} onChange={(e) => setFromServer(e.target.checked)} disabled={busy} />
          <span className="small">Load BOM from server</span>
        </label>
        <button type="button" className="btn sm primary" onClick={() => loadBom(selectedList)} disabled={busy || selected.size === 0}>Load BOM</button>
        <button type="button" className="btn sm" onClick={downloadHtml} disabled={busy || results.length === 0}>Download HTML</button>
        <button type="button" className="btn sm" onClick={fetchPdm} disabled={busy || pdmBusy}
          title={`PDM master: ${n0(pdmTotal)} row(s)${pdmFetchedAt ? ", fetched " + pdmFetchedAt : ""}`}>
          {pdmBusy ? "Fetching…" : "Fetch PDM"}
        </button>
      </div>

      <div className="t5-models">
        <div className="t5-models-head">
          <strong className="small">Cached models</strong>
          <span className="br-badge">{n0(selected.size)} selected</span>
          <span className="muted small">{n0(shownModels.length)} / {n0(cached.length)} shown</span>
          <button type="button" className="br-linkbtn" onClick={() => setSelected((cur) => new Set([...cur, ...shownModels.map((m) => m.name)]))} disabled={busy}>All</button>
          <button type="button" className="br-linkbtn" onClick={() => setSelected(new Set())} disabled={busy}>None</button>
        </div>
        {cached.length === 0 ? (
          <div className="muted small">No successfully cached BOM model yet. Type a similar model name above to load the first one.</div>
        ) : shownModels.length === 0 ? (
          <div className="muted small">No model matches "{filter}".</div>
        ) : (
          <div className="t5-chips">
            {shownModels.map((m) => {
              const on = [...selected].some((s) => s.toUpperCase() === m.name.toUpperCase());
              return (
                <label key={m.name} className={"t5-chip" + (on ? " selected" : "")} title={m.groups}>
                  <input type="checkbox" checked={on} onChange={(e) => toggleModel(m.name, e.target.checked)} disabled={busy} />
                  <span>{m.name}</span>
                </label>
              );
            })}
          </div>
        )}
      </div>

      {pdmStatus && <div className={"small t5-line " + (pdmError ? "err" : "muted")}>{pdmStatus}</div>}

      {busy ? (
        <div className="br-alert info"><span className="br-spin" /> {status || "Working..."}</div>
      ) : status ? (
        <div className="br-alert ok">{status}</div>
      ) : null}

      {costBusy ? (
        <div className="br-alert info"><span className="br-spin" /> Looking up material unit prices...</div>
      ) : costError ? (
        <div className="br-alert warn">{costError}</div>
      ) : null}

      {results.map((result, idx) => {
        const rows = visible[idx] ?? [];
        const matched = matchedProducts(result);
        return (
          <section key={result.model} className="t5-card">
            <header className="t5-card-head">
              <strong>{result.model}</strong>
              <span className="br-badge">{n0(rows.length)} rows</span>
              {hideNoise && result.rows.length !== rows.length && <span className="muted small">{n0(result.rows.length - rows.length)} hidden</span>}
              {matched.length > 0 && <span className="muted small" title="Products matched in MATE for this model name">matched: {matched.join(", ")}</span>}
              {result.loadSource && <span className={"br-badge " + (result.loadSource === "cache" ? "ok" : "info")} title={result.fetchedAt}>{result.loadSource}</span>}
              {result.warning && <span className="br-badge warn" title={result.warning}>cache warning</span>}
              {result.error && <span className="br-badge bad">{result.error}</span>}
            </header>
            {rows.length > 0 ? (
              <div className="t5-table-wrap">
                <table className="t5-table">
                  <colgroup>
                    <col className="c-level" /><col className="c-material" /><col className="c-code" /><col className="c-price" />
                    <col className="c-cur" /><col className="c-vnd" /><col className="c-pdm" /><col className="c-type" />
                  </colgroup>
                  <thead>
                    <tr>
                      <th>Lv</th>
                      <th>Material</th>
                      <th>Code</th>
                      <th className="num" title="Purchasing info record (INFR) first, else goods-receipt amount ÷ quantity (STBI)">Unit Price</th>
                      <th>Cur</th>
                      <th className="num" title="VND unit price at the month-start rate of TCURR (KURST=BWCU)">VND</th>
                      <th>PDMNO</th>
                      <th>PDTNO_TX</th>
                    </tr>
                  </thead>
                  <tbody>
                    {byProduct(rows, hideNoise).map((product) => [
                      <tr key={"p:" + product.product} className={"t5-product" + (product.isShared ? " shared" : "")}>
                        <td colSpan={8}>{product.product} <span className="br-badge light">{n0(product.rows.length)}</span></td>
                      </tr>,
                      ...product.rows.map((row) => {
                        const p = findPdm(pdm, row);
                        const cost = costs.get(row.materialCode.trim().toUpperCase());
                        return (
                          <tr key={product.product + "|" + row.productCode + "|" + row.materialCode + "|" + row.bomPathText}
                            className={"t5-lv" + Math.min(Math.max(row.bomLevel, 1), 6)}>
                            <td className="muted">{row.bomLevel}</td>
                            <td title={row.bomPathText}>
                              <span className="t5-tree" style={{ paddingLeft: (Math.max(row.bomLevel, 1) - 1) * 18 }}>
                                <span className="t5-marker">{row.hasChildren ? "▾" : "·"}</span>
                                <span>{displayName(row)}</span>
                              </span>
                            </td>
                            <td className="muted small">{row.materialCode}</td>
                            <td className="num small" title={cost?.priceSource ? `Price source ${cost.priceSource}` : undefined}>{formatUnitPrice(cost?.unitPrice)}</td>
                            <td className="muted small">{cost?.priceCurrency ?? ""}</td>
                            <td className="num small">{formatVnd(cost?.unitPriceVnd)}</td>
                            <td className="small">
                              {p?.pdmno.trim() && (
                                <button type="button" className="br-linkbtn" onClick={() => downloadDrawing(p.pdmno, p.pdmvr)}
                                  disabled={downloadBusy} title="Download the DWG for this drawing">{p.pdmno}</button>
                              )}
                            </td>
                            <td className="muted small">{p?.pdtnoTx ?? ""}</td>
                          </tr>
                        );
                      }),
                    ])}
                  </tbody>
                </table>
              </div>
            ) : !result.error.trim() ? (
              <div className="muted small t5-empty">No BOM rows found for this model.</div>
            ) : null}
          </section>
        );
      })}
    </div>
  );
}
