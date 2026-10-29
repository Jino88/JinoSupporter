import { useRef, useState } from "react";
import { toolSend } from "../toolsApi";
import {
  exampleData,
  normalizeGraphData,
  parseTable,
  type Table,
} from "./graphMaker/data";
import {
  buildRequests,
  exampleOptions,
  graphDefaults,
  graphKinds,
  loadedOptions,
  type GraphOptions as Options,
} from "./graphMaker/options";
import GraphOptions from "./graphMaker/GraphOptions";
import DataSheet from "./graphMaker/DataSheet";
import Charts, { type Payload } from "./graphMaker/Charts";
import "../tools.css";

interface Example {
  text: string;
  name: string;
  headerNotUse: boolean;
}
export default function GraphMaker() {
  const [text, setText] = useState("");
  const [delimiter, setDelimiter] = useState("auto");
  const [header, setHeader] = useState(1);
  const [kind, setKind] = useState(0);
  const [options, setOptions] = useState<Options>({ ...graphDefaults });
  const [table, setTable] = useState<Table>({ headers: [], rows: [] });
  const [payloads, setPayloads] = useState<Payload[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("Ready");
  const [source, setSource] = useState("Excel paste");
  const scatterExampleIndex = useRef(0);
  const hasData = table.headers.length > 0 && table.rows.length > 0;

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
      setSource("Excel paste");
      setStatus(parsed.rows.length ? "Ready to render" : "No rows found.");
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
    setOptions((current) => ({
      ...current,
      xColumn: "",
      xColumns: [],
      yColumns: [],
      graphColumnGroups: {},
      heatMapXColumn: "",
      heatMapYColumn: "",
      heatMapValueColumn: "",
      heatMapNgColumn: "",
      heatMapInputColumn: "",
      graphSpecColumn: "",
      graphUpperColumn: "",
      graphLowerColumn: "",
      scatterHeaderNotUse: false,
    }));
    setPayloads([]);
    setStatus("Ready");
    setError("");
  }
  async function renderCharts(
    data = text,
    settings = options,
    separator = delimiter,
    headerRow = header,
  ) {
    setBusy(true);
    setError("");
    try {
      const results: Payload[] = [];
      for (const request of buildRequests(kind, settings)) {
        results.push(
          await toolSend<Payload>("graph-maker/render", {
            text: data,
            delimiter: separator,
            headerRowNumber: headerRow,
            headerNotUse: kind === 1 && settings.scatterHeaderNotUse,
            request,
          }),
        );
      }
      setPayloads(results);
      const normalizedKind = [
        "Line",
        "Scatter",
        "Average",
        "Cpk",
        "NormalDistribution",
        "NoXMultiY",
        "ProcessTrend",
        "HeatMap",
      ].indexOf(results[0]?.kind);
      if (normalizedKind >= 0) setKind(normalizedKind);
      setStatus(`${results.length} graph(s) rendered`);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function loadExample() {
    setBusy(true);
    setError("");
    try {
      const index = kind === 1 ? scatterExampleIndex.current++ : 0;
      const example = await toolSend<Example>(
        `graph-maker/example/${kind}/${index}`,
        undefined,
        "GET",
      );
      const data = exampleData(example.text);
      const parsed = parseTable(data, "auto", 1, example.headerNotUse);
      const settings = exampleOptions(kind, example.name, parsed, {
        ...options,
        scatterHeaderNotUse: example.headerNotUse,
      });
      setDelimiter("auto");
      setHeader(1);
      setText(data);
      setTable(parsed);
      setOptions(settings);
      setSource(example.name);
      setStatus("Example loaded");
      await renderCharts(data, settings, "auto", 1);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  }
  async function loadFile(file?: File) {
    if (!file) return;
    try {
      pasteData(await file.text());
      setSource(file.name);
    } catch (e) {
      setError((e as Error).message);
    }
  }

  return (
    <div className="tools-page">
      <header className="page-head">
        <h1>Graph Maker</h1>
        <p className="muted">
          One workspace for line, scatter, average, CPK, distribution, heatmap,
          and process trend views.
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
              className="btn sm primary"
              disabled={!hasData}
              onClick={() => void renderCharts()}
            >
              Render
            </button>
            <button className="btn sm" onClick={() => void loadExample()}>
              Example
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
                {busy ? "Working…" : status}{" "}
                {hasData &&
                  ` · ${source} · ${table.rows.length.toLocaleString("en-US")} rows · ${table.headers.length} columns`}
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
                <summary>Text and file input</summary>
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
                <label className="tool-field">
                  Load data file
                  <input
                    type="file"
                    accept=".txt,.csv,.tsv,.json"
                    onChange={(e) => void loadFile(e.target.files?.[0])}
                  />
                </label>
              </details>
            </div>
            <Charts payloads={payloads} onError={setError} />
          </section>
        </div>
      </fieldset>
    </div>
  );
}
