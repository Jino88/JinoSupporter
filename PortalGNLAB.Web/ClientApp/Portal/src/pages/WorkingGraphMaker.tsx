import { useEffect, useRef, useState } from "react";
import { normalizeGraphData, parseTable, type Table } from "./graphMaker/data";
import {
  graphDefaults,
  graphKinds,
  loadedOptions,
  type GraphOptions as Options,
} from "./graphMaker/options";
import GraphOptions from "./graphMaker/GraphOptions";
import DataSheet from "./graphMaker/DataSheet";
import Charts, { type Payload } from "./graphMaker/Charts";
import {
  analyzeAndRender,
  renderGraphs,
  type AnalyzeResult,
} from "./workingGraph/analyze";
import "../tools.css";

export default function WorkingGraphMaker() {
  const [text, setText] = useState("");
  const [delimiter, setDelimiter] = useState("auto");
  const [header, setHeader] = useState(1);
  const [kind, setKind] = useState(0);
  const [options, setOptions] = useState<Options>({ ...graphDefaults });
  const [table, setTable] = useState<Table>({ headers: [], rows: [] });
  const [payloads, setPayloads] = useState<Payload[]>([]);
  const [analysis, setAnalysis] = useState<AnalyzeResult | null>(null);
  const [analyzing, setAnalyzing] = useState(false);
  const [rendering, setRendering] = useState(false);
  const [elapsed, setElapsed] = useState(0);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Ready");
  const abort = useRef<AbortController | null>(null);
  const hasData = table.headers.length > 0 && table.rows.length > 0;
  const busy = analyzing || rendering;
  const noHeader = kind === 1 && options.scatterHeaderNotUse;

  useEffect(() => {
    if (!analyzing) return;
    const started = Date.now();
    setElapsed(0);
    const timer = window.setInterval(
      () => setElapsed(Math.floor((Date.now() - started) / 1000)),
      1000,
    );
    return () => window.clearInterval(timer);
  }, [analyzing]);
  useEffect(() => () => abort.current?.abort(), []);

  function pasteData(
    value: string,
    settings = options,
    separator = delimiter,
    headerRow = header,
  ) {
    try {
      const normalized = normalizeGraphData(value);
      const parsed = parseTable(
        normalized,
        separator,
        headerRow,
        kind === 1 && settings.scatterHeaderNotUse,
      );
      setText(normalized);
      setTable(parsed);
      setOptions(loadedOptions(parsed, settings));
      setPayloads([]);
      setAnalysis(null);
      setStatus(parsed.rows.length ? "Ready to analyze" : "No rows found.");
      setError("");
    } catch (e) {
      setError((e as Error).message);
    }
  }
  function changeNoHeader(enabled: boolean) {
    const next = { ...options, scatterHeaderNotUse: enabled };
    setOptions(next);
    if (text.trim()) pasteData(text, next);
  }
  function clearTable() {
    setText("");
    setTable({ headers: [], rows: [] });
    setOptions({ ...graphDefaults });
    setPayloads([]);
    setAnalysis(null);
    setStatus("Ready");
    setError("");
  }
  async function analyze() {
    const controller = new AbortController();
    abort.current = controller;
    setAnalyzing(true);
    setError("");
    setStatus("AI is analyzing the data…");
    try {
      const done = await analyzeAndRender(
        { text, delimiter, headerRowNumber: header, headerNotUse: noHeader },
        table,
        controller.signal,
        (r, o) => {
          setAnalysis(r);
          setKind(r.kind);
          setOptions(o);
          setStatus("Analysis done, rendering…");
          setAnalyzing(false);
          setRendering(true);
        },
      );
      setPayloads(done.payloads);
      setStatus(`${done.payloads.length} graph(s) rendered`);
    } catch (e) {
      if (controller.signal.aborted) setStatus("Analysis cancelled");
      else setError((e as Error).message);
    } finally {
      setAnalyzing(false);
      setRendering(false);
    }
  }
  async function render() {
    setRendering(true);
    setError("");
    try {
      const drawn = await renderGraphs(
        { text, delimiter, headerRowNumber: header, headerNotUse: noHeader },
        kind,
        options,
      );
      setPayloads(drawn);
      setStatus(`${drawn.length} graph(s) rendered`);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setRendering(false);
    }
  }

  return (
    <div className="tools-page">
      <header className="page-head">
        <h1>Working · Graphmaker</h1>
        <p className="muted">
          Paste your data and press AI analyze: the AI picks the graph type and
          options, then draws it. You can adjust the result and render again.
        </p>
      </header>
      {error && (
        <div className="alert" role="alert">
          {error}
        </div>
      )}
      <fieldset className="tool-fieldset" disabled={busy}>
        <section className="card">
          <div className="form-row">
            <button
              className="btn sm primary"
              disabled={!hasData}
              onClick={() => void analyze()}
            >
              AI analyze
            </button>
            <label>
              Graph type
              <select
                className="inp"
                value={kind}
                onChange={(e) => setKind(Number(e.target.value))}
              >
                {graphKinds.map((name, index) => (
                  <option key={index} value={index}>
                    {name}
                  </option>
                ))}
              </select>
            </label>
            <button
              className="btn sm"
              disabled={!hasData}
              onClick={() => void render()}
            >
              Render
            </button>
            <label>
              Delimiter
              <select
                className="inp"
                value={delimiter}
                onChange={(e) => setDelimiter(e.target.value)}
              >
                {[
                  ["auto", "Auto"],
                  [",", "Comma"],
                  ["\t", "Tab"],
                  [";", "Semicolon"],
                  ["|", "Pipe"],
                ].map(([value, label]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
            <label>
              Header row
              <input
                className="inp"
                type="number"
                min={1}
                value={header}
                onChange={(e) => setHeader(Math.max(1, Number(e.target.value)))}
              />
            </label>
            <button className="btn sm" disabled={!hasData} onClick={clearTable}>
              Table Clear
            </button>
          </div>
        </section>
      </fieldset>
      {analyzing && (
        <section className="card" role="status">
          <p>
            AI is analyzing the data… {elapsed}s elapsed (usually 30 s to 3
            min)
          </p>
          <button className="btn sm" onClick={() => abort.current?.abort()}>
            Cancel
          </button>
        </section>
      )}
      {analysis && <AnalysisCard analysis={analysis} />}
      <fieldset className="tool-fieldset" disabled={busy}>
        <div className="graph-layout">
          <GraphOptions
            kind={kind}
            headers={table.headers}
            options={options}
            onChange={setOptions}
            onNoHeader={changeNoHeader}
          />
          <section className="graph-workspace">
            <div className="card">
              <p role="status">
                {rendering ? "Rendering…" : status}
                {hasData &&
                  ` · ${table.rows.length.toLocaleString("en-US")} rows · ${table.headers.length} columns`}
              </p>
              <p className="muted small">
                Click the data table and press Ctrl+V to paste Excel cells.
              </p>
              <DataSheet
                table={table}
                kind={kind}
                options={options}
                onPaste={pasteData}
              />
              <details>
                <summary>Text input</summary>
                <label className="tool-field">
                  Delimited data
                  <textarea
                    className="inp tool-textarea"
                    value={text}
                    onChange={(e) => setText(e.target.value)}
                  />
                </label>
                <button className="btn sm" onClick={() => pasteData(text)}>
                  Load data
                </button>
              </details>
            </div>
            <Charts payloads={payloads} onError={setError} />
          </section>
        </div>
      </fieldset>
    </div>
  );
}

export function AnalysisCard({ analysis }: { analysis: AnalyzeResult }) {
  return (
    <section className="card">
      <h2>AI pick: {graphKinds[analysis.kind] ?? `Type ${analysis.kind}`}</h2>
      {analysis.title && (
        <p>
          <strong>{analysis.title}</strong>
        </p>
      )}
      <p>{analysis.reason}</p>
      <p className="muted small">
        Model {analysis.model} · {analysis.elapsedSeconds}s
      </p>
    </section>
  );
}
