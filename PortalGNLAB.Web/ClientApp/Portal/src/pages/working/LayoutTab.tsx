import { useEffect, useMemo, useRef, useState } from "react";
import { useTr } from "./i18n";
import { deleteLayout, fitView, loadLayout, panBy, textShown, toggleLayer, toScreen, uploadLayout, wheelFactor, zoomAt } from "./layout";
import type { Layout, View } from "./layout";

const stamp = (iso: string) => iso.slice(0, 16).replace("T", " ");
const megabytes = (n: number) => (n / 1048576).toFixed(1) + " MB";

/** Path bounding boxes, computed once so panning can skip what is off screen. */
function boxes(layout: Layout): Float64Array {
  const out = new Float64Array(layout.paths.length * 4);
  layout.paths.forEach((path, i) => {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    const p = path.p;
    for (let k = 0; k + 1 < p.length; k += 2) {
      if (p[k] < x0) x0 = p[k];
      if (p[k] > x1) x1 = p[k];
      if (p[k + 1] < y0) y0 = p[k + 1];
      if (p[k + 1] > y1) y1 = p[k + 1];
    }
    out.set([x0, y0, x1, y1], i * 4);
  });
  return out;
}

function draw(canvas: HTMLCanvasElement, layout: Layout, box: Float64Array, view: View, hidden: ReadonlySet<number>) {
  const ctx = canvas.getContext("2d");
  if (!ctx) return;
  const dpr = window.devicePixelRatio || 1;
  const w = canvas.width / dpr, h = canvas.height / dpr;
  const fg = getComputedStyle(canvas).color || "#000";
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  ctx.lineWidth = 1;
  ctx.lineJoin = "round";
  // World window on screen, to skip what is outside it.
  const wx0 = (0 - view.tx) / view.scale, wx1 = (w - view.tx) / view.scale;
  const wy0 = -(h - view.ty) / view.scale, wy1 = -(0 - view.ty) / view.scale;

  const byColor = new Map<string, number[]>();
  layout.paths.forEach((path, i) => {
    if (hidden.has(path.l) || path.p.length < 4) return;
    if (box[i * 4 + 2] < wx0 || box[i * 4] > wx1 || box[i * 4 + 3] < wy0 || box[i * 4 + 1] > wy1) return;
    const c = path.c ?? fg;
    const list = byColor.get(c);
    if (list) list.push(i); else byColor.set(c, [i]);
  });
  for (const [color, list] of byColor) {
    ctx.strokeStyle = color;
    ctx.beginPath();
    for (const i of list) {
      const { p, z } = layout.paths[i];
      ctx.moveTo(p[0] * view.scale + view.tx, -p[1] * view.scale + view.ty);
      for (let k = 2; k + 1 < p.length; k += 2) ctx.lineTo(p[k] * view.scale + view.tx, -p[k + 1] * view.scale + view.ty);
      if (z) ctx.closePath();
    }
    ctx.stroke();
  }

  const baselines = ["alphabetic", "bottom", "middle", "top"] as const;
  const aligns = ["left", "center", "right"] as const;
  for (const t of layout.texts) {
    if (hidden.has(t.l) || !t.t || !textShown(view, t.h)) continue;
    const [sx, sy] = toScreen(view, t.x, t.y);
    const px = t.h * view.scale;
    if (sx < -px * 50 || sx > w + px * 50 || sy < -px * 50 || sy > h + px * 50) continue;
    ctx.fillStyle = t.c ?? fg;
    ctx.font = `${px}px sans-serif`;
    ctx.textAlign = aligns[t.a ?? 0];
    ctx.textBaseline = baselines[t.v ?? 0];
    ctx.save();
    ctx.translate(sx, sy);
    if (t.r) ctx.rotate((-t.r * Math.PI) / 180);
    t.t.split("\n").forEach((line, n) => ctx.fillText(line, 0, n * px * 1.2));
    ctx.restore();
  }
}

/** Canvas viewer: wheel zooms around the cursor, dragging pans, 1px lines, redrawn once per frame. */
function Viewer(props: { layout: Layout; hidden: ReadonlySet<number>; fitSignal: number }) {
  const { layout, hidden } = props;
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const viewRef = useRef<View | null>(null);
  const hiddenRef = useRef(hidden);
  const frame = useRef(0);
  const box = useMemo(() => boxes(layout), [layout]);
  const paint = useRef<() => void>(() => {});

  hiddenRef.current = hidden;
  paint.current = () => {
    cancelAnimationFrame(frame.current);
    frame.current = requestAnimationFrame(() => {
      const canvas = canvasRef.current;
      if (canvas && viewRef.current) draw(canvas, layout, box, viewRef.current, hiddenRef.current);
    });
  };

  const fit = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const dpr = window.devicePixelRatio || 1;
    viewRef.current = fitView(layout.bounds, canvas.width / dpr, canvas.height / dpr);
    paint.current();
  };

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    viewRef.current = null;
    const resize = () => {
      const dpr = window.devicePixelRatio || 1;
      const { width, height } = canvas.getBoundingClientRect();
      if (width < 1 || height < 1) return;
      canvas.width = Math.round(width * dpr);
      canvas.height = Math.round(height * dpr);
      if (!viewRef.current) fit(); else paint.current();
    };
    const observer = new ResizeObserver(resize);
    observer.observe(canvas);
    const wheel = (e: WheelEvent) => {
      e.preventDefault();
      const view = viewRef.current;
      if (!view) return;
      const r = canvas.getBoundingClientRect();
      viewRef.current = zoomAt(view, e.clientX - r.left, e.clientY - r.top, wheelFactor(e.deltaY));
      paint.current();
    };
    canvas.addEventListener("wheel", wheel, { passive: false });
    resize();
    return () => { observer.disconnect(); canvas.removeEventListener("wheel", wheel); cancelAnimationFrame(frame.current); };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [layout]);

  useEffect(() => { paint.current(); }, [hidden]);
  useEffect(() => {
    if (props.fitSignal) fit();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.fitSignal]);

  const drag = useRef<{ x: number; y: number } | null>(null);
  return (
    <canvas ref={canvasRef} className="wk-lay-canvas"
      onPointerDown={(e) => { drag.current = { x: e.clientX, y: e.clientY }; e.currentTarget.setPointerCapture(e.pointerId); }}
      onPointerMove={(e) => {
        const d = drag.current;
        if (!d || !viewRef.current) return;
        viewRef.current = panBy(viewRef.current, e.clientX - d.x, e.clientY - d.y);
        drag.current = { x: e.clientX, y: e.clientY };
        paint.current();
      }}
      onPointerUp={() => { drag.current = null; }}
      onPointerCancel={() => { drag.current = null; }} />
  );
}

/** The model's process layout (a DXF read by the server), drawn on a canvas. */
export default function LayoutTab(props: {
  modelName: string;
  onError: (e: unknown) => void;
  /** Layout already read (tests): null = none saved; else read here. */
  initial?: Layout | null;
}) {
  const tr = useTr();
  const [layout, setLayout] = useState<Layout | null | undefined>(props.initial);
  const [hidden, setHidden] = useState<ReadonlySet<number>>(new Set());
  const [fitSignal, setFitSignal] = useState(0);
  const [busy, setBusy] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (props.initial !== undefined) return;
    let live = true;
    loadLayout(props.modelName).then((l) => { if (live) setLayout(l); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const send = async (file: File | undefined) => {
    if (!file) return;
    setBusy(true);
    try {
      setLayout(await uploadLayout(props.modelName, file));
      setHidden(new Set());
    } catch (e) { props.onError(e); } finally { setBusy(false); if (input.current) input.current.value = ""; }
  };
  const remove = async () => {
    if (!window.confirm(tr("Delete the process layout of this model?"))) return;
    setBusy(true);
    try { await deleteLayout(props.modelName); setLayout(null); setHidden(new Set()); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };

  const picker = <input ref={input} type="file" accept=".dxf" hidden disabled={busy} onChange={(e) => send(e.target.files?.[0])} />;
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>Process Layout</h2>
        {layout && (
          <span className="muted small">
            <b>{layout.fileName}</b> · {megabytes(layout.size)} · {stamp(layout.uploadedAt)} · {layout.uploadedBy}
          </span>
        )}
      </div>
      <div className="pi-sec-b">
        {layout === undefined ? <div className="muted">{tr("Loading…")}</div>
          : layout === null ? (
            <div className="wk-lay-empty">
              <div className="muted">{tr("No process layout is registered for this model. Upload its DXF drawing to see it here.")}</div>
              <button className="btn sm primary" disabled={busy} onClick={() => input.current?.click()}>{busy ? tr("Reading the drawing…") : tr("Upload DXF")}</button>
              {picker}
            </div>
          ) : (
            <>
              <div className="wk-lay-bar">
                <button className="btn sm" onClick={() => setFitSignal((n) => n + 1)}>{tr("Fit")}</button>
                <button className="btn sm" disabled={busy} onClick={() => input.current?.click()}>{busy ? tr("Reading the drawing…") : tr("Replace")}</button>
                <button className="btn sm" disabled={busy} onClick={remove}>{tr("Delete")}</button>
                {picker}
                <span className="wk-lay-layers">
                  {layout.layers.map((name, i) => (
                    <button key={i} className={"wk-chip" + (hidden.has(i) ? "" : " on")} title={name}
                      onClick={() => setHidden(toggleLayer(hidden, i))}>{name}</button>
                  ))}
                </span>
              </div>
              <Viewer layout={layout} hidden={hidden} fitSignal={fitSignal} />
            </>
          )}
      </div>
    </section>
  );
}
