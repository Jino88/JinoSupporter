import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { qrApi, type GaussScatter } from "../qrApi";
import { DATE_SETTLE_MS, useQrPrefs, useSettleTimer } from "./qr/hooks";
import {
  MAX_ROW_CHOICES,
  PLOT,
  CURVE_BASE,
  VALUE2_COLOR,
  VALUE_COLOR,
  addDays,
  axisNum,
  buildModelOptions,
  buildScatter,
  diamond,
  formatPpm,
  localToday,
  mapY,
  markNum,
  matchesModel,
  maxRowOptions,
  nudged,
  num,
  num0,
  parseDay,
  pointX,
  validModel,
} from "./qr/logic";
import "../qr.css";

const f = (value: number) => String(Math.round(value * 100) / 100);

function Spinner() {
  return <span className="gsc-spinner" aria-hidden="true" />;
}

/** /app/qr/gauss-scatter — port of QrGaussScatterPage.razor. ?from= &to= &max= &series=1|2 come from QR SUMMARY. */
export default function QrGaussScatter() {
  const { prefs, error: prefsError } = useQrPrefs();
  const initial = useMemo(() => {
    const params = new URLSearchParams(window.location.search);
    const max = Number.parseInt(params.get("max") ?? "", 10);
    const series = params.get("series");
    return {
      from: parseDay(params.get("from")),
      to: parseDay(params.get("to")),
      max: max > 0 ? max : 0,
      showValue: series !== "2",
      showValue2: series !== "1",
    };
  }, []);

  const [result, setResult] = useState<GaussScatter | null>(null);
  const [from, setFrom] = useState(initial.from);
  const [to, setTo] = useState(initial.to);
  const [maxRows, setMaxRows] = useState(initial.max || 2000);
  const [modelCode, setModelCode] = useState("");
  const [showValue, setShowValue] = useState(initial.showValue);
  const [showValue2, setShowValue2] = useState(initial.showValue2);
  const [ngOnly, setNgOnly] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const settle = useSettleTimer();
  const loading = useRef(false);

  const load = useCallback(async (fromV: string, toV: string, max: number) => {
    if (loading.current) return;
    loading.current = true;
    setBusy(true);
    setError("");
    try {
      const next = await qrApi.scatter(fromV, toV, max);
      setResult(next);
      setFrom(next.from);
      setTo(next.to);
      setMaxRows(next.maxRows);
    } catch (e) {
      setError((e as Error).message);
    } finally {
      loading.current = false;
      setBusy(false);
    }
  }, []);

  // The window is opened from QR SUMMARY with ?max=, which wins; a direct visit uses the cap last picked.
  useEffect(() => {
    if (!prefs?.configured) return;
    const max = initial.max || (prefs.maxQr > 0 ? prefs.maxQr : 2000);
    setMaxRows(max);
    void load(initial.from, initial.to, max);
  }, [prefs, load, initial]);

  const models = useMemo(() => buildModelOptions(prefs?.modelCodes ?? [], result?.productIds ?? []), [prefs, result]);
  const activeModel = validModel(modelCode, models.options);

  // Re-filter the points already fetched: picking a model costs no query.
  const points = useMemo(() => {
    const keep = (p: { productId: string }) => matchesModel(p.productId, activeModel, models.known);
    return { value: (result?.value ?? []).filter(keep), value2: (result?.value2 ?? []).filter(keep) };
  }, [result, activeModel, models.known]);

  const model = useMemo(
    () => buildScatter({ value: points.value, value2: points.value2, showValue, showValue2, ngOnly }),
    [points, showValue, showValue2, ngOnly],
  );
  const { views, scale, columns } = model;

  /** Toggles one reading; the last visible one stays on so the chart is never empty. */
  const toggleSeries = (slot: 0 | 1) => {
    if (slot === 0) {
      if (showValue && !showValue2) return;
      setShowValue(!showValue);
    } else {
      if (showValue2 && !showValue) return;
      setShowValue2(!showValue2);
    }
  };

  const queueReload = (fromV: string, toV: string) => settle.schedule(() => void load(fromV, toV, maxRows), DATE_SETTLE_MS);
  const onFrom = (value: string) => {
    if (!value) return;
    setFrom(value);
    queueReload(value, to);
  };
  const onTo = (value: string) => {
    if (!value) return;
    setTo(value);
    queueReload(from, value);
  };
  const quickRange = (daysBack: number) => {
    settle.cancel();
    const anchor = result?.maxDate ?? localToday();
    setTo(anchor);
    setFrom(addDays(anchor, -daysBack));
    void load(addDays(anchor, -daysBack), anchor, maxRows);
  };
  const onMax = (value: number) => {
    if (!(value > 0)) return;
    setMaxRows(value);
    void qrApi.savePrefs({ maxQr: value }).catch(() => undefined);
    void load(from, to, value);
  };

  return (
    <div className="gsc-page">
      <header className="gsc-titlebar">
        <div>
          <div className="gsc-kicker">GAUSS · dbo.GSTD · DISPERSION</div>
          <h1>GAUSS SCATTER</h1>
          <p>One column per test date. Value and Value 2 are drawn together in their own colours, each with its own normal curve and box.</p>
        </div>
        <div className="gsc-actions">
          <div className="gsc-quick">
            <button className={`gsc-chip gsc-swatch-chip ${showValue ? "gsc-on" : ""}`} type="button" onClick={() => toggleSeries(0)}>
              <i className="gsc-dot" style={{ background: VALUE_COLOR }} />Value
            </button>
            <button className={`gsc-chip gsc-swatch-chip ${showValue2 ? "gsc-on" : ""}`} type="button" onClick={() => toggleSeries(1)}>
              <i className="gsc-dot" style={{ background: VALUE2_COLOR }} />Value 2
            </button>
          </div>
          <div className="gsc-quick">
            <button className={`gsc-chip ${ngOnly ? "" : "gsc-on"}`} type="button" onClick={() => setNgOnly(false)}>All</button>
            <button className={`gsc-chip ${ngOnly ? "gsc-on" : ""}`} type="button" onClick={() => setNgOnly(true)}>NG only</button>
          </div>
        </div>
      </header>

      <div className="gsc-toolbar">
        <label className="gsc-field">
          <span>From</span>
          <input type="date" value={from} onChange={(e) => onFrom(e.target.value)} />
        </label>
        <label className="gsc-field">
          <span>To</span>
          <input type="date" value={to} onChange={(e) => onTo(e.target.value)} />
        </label>
        <div className="gsc-quick">
          <button className="gsc-chip" type="button" disabled={busy} onClick={() => quickRange(0)}>Latest day</button>
          <button className="gsc-chip" type="button" disabled={busy} onClick={() => quickRange(1)}>2 days</button>
          <button className="gsc-chip" type="button" disabled={busy} onClick={() => quickRange(6)}>7 days</button>
          <button className="gsc-chip" type="button" disabled={busy} onClick={() => quickRange(29)}>30 days</button>
        </div>
        <label className="gsc-field">
          <span>Max QR</span>
          <select disabled={busy} value={maxRows} onChange={(e) => onMax(Number(e.target.value))}>
            {maxRowOptions(MAX_ROW_CHOICES, maxRows).map((option) => (
              <option key={option} value={option}>{num0(option)}</option>
            ))}
          </select>
        </label>
        <label className="gsc-field">
          <span>Model</span>
          <select disabled={busy} value={activeModel} onChange={(e) => setModelCode(e.target.value)}>
            <option value="">All models</option>
            {models.options.map((option) => (
              <option key={option.code} value={option.code}>{option.label}</option>
            ))}
          </select>
        </label>
        <button className="gsc-btn gsc-primary" type="button" onClick={() => { settle.cancel(); void load(from, to, maxRows); }} disabled={busy}>
          {busy ? (<><Spinner /><span>Loading</span></>) : <span>Refresh</span>}
        </button>
      </div>

      {(error || prefsError) && <div className="gsc-alert gsc-danger">{error || prefsError}</div>}
      {result?.warning && <div className="gsc-alert gsc-warning">{result.warning}</div>}

      {prefs && !prefs.configured ? (
        <div className="gsc-alert gsc-warning">QR DB password is not configured on the server.</div>
      ) : busy && !result ? (
        <div className="gsc-state"><Spinner /><span>Reading GSTD from TCDB.</span></div>
      ) : !result ? null : model.totalPoints === 0 ? (
        <div className="gsc-state">No numeric gauss reading in {result.from} ~ {result.to}.</div>
      ) : (
        <>
          <div className="gsc-stats">
            <table>
              <thead>
                <tr>
                  <th>Series</th><th>Points</th><th>Mean</th><th>σ</th><th>Median</th>
                  <th>Q1</th><th>Q3</th><th>Min</th><th>Max</th><th>Range</th><th>IQR</th><th>NG</th>
                </tr>
              </thead>
              <tbody>
                {views.map((view) => (
                  <tr key={view.slot}>
                    <th scope="row"><i className="gsc-dot" style={{ background: view.color }} />{view.label}</th>
                    <td>{num0(view.stats.n)}</td>
                    <td>{num(view.stats.mean)}</td>
                    <td>{num(view.stats.stdDev)}</td>
                    <td>{num(view.stats.median)}</td>
                    <td>{num(view.stats.q1)}</td>
                    <td>{num(view.stats.q3)}</td>
                    <td>{num(view.stats.min)}</td>
                    <td>{num(view.stats.max)}</td>
                    <td>{num(view.stats.max - view.stats.min)}</td>
                    <td>{num(view.stats.q3 - view.stats.q1)}</td>
                    <td>{num0(view.ngCount)} · {formatPpm(view.ngRatePpm)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            <span className="gsc-stats-note">{num0(model.dates.length)} test date(s) · {result.from} ~ {result.to}</span>
          </div>

          <div className="gsc-plot">
            <svg
              viewBox="0 0 1320 566"
              preserveAspectRatio="xMidYMid meet"
              role="img"
              aria-label="GAUSS Value and Value 2 by test date, each with its fitted normal curve and a box labelled with max, Q3, Q2, mean, Q1, min and IQR"
            >
              {/* Y grid, shared by both panels */}
              {scale.ticks.map((tick) => (
                <g key={tick}>
                  <line x1={74} y1={f(mapY(scale, tick))} x2={890} y2={f(mapY(scale, tick))} stroke="#e4e8ee" strokeWidth={1} />
                  <line x1={930} y1={f(mapY(scale, tick))} x2={1300} y2={f(mapY(scale, tick))} stroke="#e4e8ee" strokeWidth={1} />
                  <text x={68} y={f(mapY(scale, tick) + 3.4)} textAnchor="end" className="gsc-axis">{axisNum(tick)}</text>
                </g>
              ))}

              <rect x={74} y={34} width={816} height={466} fill="none" stroke="#c6ccd5" strokeWidth={1} />
              <rect x={930} y={34} width={370} height={466} fill="none" stroke="#c6ccd5" strokeWidth={1} />
              <text x={482} y={548} textAnchor="middle" className="gsc-axis-title">Test date</text>
              <text x={1115} y={548} textAnchor="middle" className="gsc-axis-title">Normal fit · distribution</text>
              <text x={18} y={267} textAnchor="middle" className="gsc-axis-title" transform="rotate(-90 18 267)">GAUSS reading</text>

              {/* One column per date */}
              {columns.map((column) => (
                <g key={column.date}>
                  {column.index > 0 && <line x1={f(column.left)} y1={34} x2={f(column.left)} y2={500} stroke="#eef1f5" strokeWidth={1} />}
                  <line x1={f(column.center)} y1={500} x2={f(column.center)} y2={505} stroke="#c6ccd5" strokeWidth={1} />
                  {column.showLabel && (
                    <g>
                      <text x={f(column.center)} y={518} textAnchor="middle" className="gsc-axis">{column.label}</text>
                      <text x={f(column.center)} y={530} textAnchor="middle" className="gsc-axis gsc-faint">{num0(column.count)}</text>
                    </g>
                  )}
                </g>
              ))}

              {/* Each series: its mean / median guides, then its own strip per date */}
              {views.map((view) => (
                <g key={view.slot}>
                  <line x1={PLOT.left} y1={f(mapY(scale, view.stats.median))} x2={PLOT.right} y2={f(mapY(scale, view.stats.median))} stroke={view.color} strokeWidth={1.3} strokeOpacity={0.85} />
                  <line x1={PLOT.left} y1={f(mapY(scale, view.stats.mean))} x2={PLOT.right} y2={f(mapY(scale, view.stats.mean))} stroke={view.color} strokeWidth={1.3} strokeOpacity={0.85} strokeDasharray="6 4" />
                  {view.plotted.map((p, i) => (
                    <circle key={i} cx={f(pointX(model, p, view))} cy={f(mapY(scale, p.value))} r={2.3} fill={view.color} fillOpacity={0.5}>
                      {view.showTitles && <title>{`${p.productId} · ${view.label} ${num(p.value)} · ${p.time.slice(5)}`}</title>}
                    </circle>
                  ))}
                </g>
              ))}

              {/* Vertical normal curve per series; with one series its histogram sits behind it. */}
              <line x1={946} y1={34} x2={946} y2={500} stroke="#c6ccd5" strokeWidth={1} />
              {views.map((view) => (
                <g key={view.slot}>
                  {view.histogram.map((bar, i) => (
                    <rect key={i} x={CURVE_BASE} y={f(bar.top)} width={f(bar.width)} height={f(bar.height)} fill={view.color} fillOpacity={0.13} />
                  ))}
                  {view.curve.length > 0 && <polyline points={view.curve} fill="none" stroke={view.color} strokeWidth={1.8} />}
                  {view.sigmaMarks.map((mark) => (
                    <g key={mark.label}>
                      <line x1={946} y1={f(mark.y)} x2={1096} y2={f(mark.y)} stroke="#9aa6b4" strokeWidth={1} strokeDasharray="3 3" />
                      <text x={1100} y={f(mark.y + 3.4)} className="gsc-axis">{mark.label}</text>
                    </g>
                  ))}
                </g>
              ))}

              {/* One box per series: whiskers to min/max, box over Q1..Q3, median solid, mean dashed with a diamond */}
              {views.map((view) => {
                const top = mapY(scale, view.stats.max);
                const bottom = mapY(scale, view.stats.min);
                const q3 = mapY(scale, view.stats.q3);
                const q1 = mapY(scale, view.stats.q1);
                const med = mapY(scale, view.stats.median);
                const mean = mapY(scale, view.stats.mean);
                const boxTop = Math.min(q1, q3);
                const boxHeight = Math.max(Math.abs(q1 - q3), 1);
                const { boxCenter: c, boxHalf: h, capHalf: cap, color } = view;
                return (
                  <g key={view.slot}>
                    <line x1={f(c)} y1={f(top)} x2={f(c)} y2={f(bottom)} stroke={color} strokeWidth={1.2} />
                    <line x1={f(c - cap)} y1={f(top)} x2={f(c + cap)} y2={f(top)} stroke={color} strokeWidth={1.6} />
                    <line x1={f(c - cap)} y1={f(bottom)} x2={f(c + cap)} y2={f(bottom)} stroke={color} strokeWidth={1.6} />
                    <rect x={f(c - h)} y={f(boxTop)} width={f(h * 2)} height={f(boxHeight)} fill={color} fillOpacity={0.08} stroke={color} strokeWidth={1.3} />
                    <line x1={f(c - h)} y1={f(med)} x2={f(c + h)} y2={f(med)} stroke={color} strokeWidth={2.4} />
                    <line x1={f(c - h)} y1={f(mean)} x2={f(c + h)} y2={f(mean)} stroke={color} strokeWidth={2} strokeDasharray="6 4" />
                    <polygon points={diamond(c, mean)} fill={color} />
                    <text x={f(c)} y={524} textAnchor="middle" className="gsc-mark" fill={color}>{view.label}</text>
                    <text x={f(c)} y={536} textAnchor="middle" className="gsc-axis">IQR {markNum(view.stats.q3 - view.stats.q1)}</text>
                    {view.boxLabels.map((label) => (
                      <g key={label.text}>
                        {nudged(label) && (
                          <line x1={f(c + h + 1)} y1={f(label.lineY)} x2={f(view.labelX - 3)} y2={f(label.textY - 3.4)} stroke={color} strokeWidth={0.8} strokeOpacity={0.45} />
                        )}
                        <text x={f(view.labelX)} y={f(label.textY)} className="gsc-mark" fill={label.fill}>{label.text}</text>
                      </g>
                    ))}
                  </g>
                );
              })}
            </svg>
          </div>

          <footer className="gsc-foot">
            {views.map((view) => (
              <span key={view.slot} className="gsc-key"><i className="gsc-dot" style={{ background: view.color }} />{view.label}</span>
            ))}
            <span className="gsc-key"><i className="gsc-line" />Median</span>
            <span className="gsc-key"><i className="gsc-line gsc-dashed" />Mean</span>
            <span className="gsc-key"><i className="gsc-box" />Q1 – Q3, whiskers to min / max</span>
            <span className="gsc-key"><i className="gsc-curve" />Normal fit{views.length === 1 ? " over the reading histogram" : ""}</span>
            {model.plottedPoints < model.totalPoints && (
              <span>Drawing {num0(model.plottedPoints)} of {num0(model.totalPoints)} points — statistics use all of them</span>
            )}
            {result.fetchedAt && <span>as of {result.fetchedAt}</span>}
          </footer>
        </>
      )}
    </div>
  );
}
