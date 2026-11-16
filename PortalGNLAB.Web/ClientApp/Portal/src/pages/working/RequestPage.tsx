import { useEffect, useState, type ClipboardEvent, type DragEvent, type ReactNode } from "react";
import DrawingPicker from "./DrawingPicker";
import { aiPictures } from "./aiPlace";
import { useTr } from "./i18n";
import {
  MAX_REQUEST_PICTURE_BYTES, METHODS, REQUEST_KINDS, blankItem, blankRequest, checkOf, commonModel, dotDate, drawingApi, drawingKey, matchText, orderOf, otherVersions, readRequest,
  requestApi, requestFields, requestFileName, requestProblem, requestTitle, totals, uncheckedNos, withModel,
  type DrawingCheck, type RequestFields, type RequestItem, type WorkingRequest,
} from "./request";
import { Thumbs, dataUrl } from "./SetupAi";
import { downloadBlob } from "../../toolsApi";
import "./request.css";

/**
 * Working > Request: PROCESSING REQUEST forms (JIG/MC). A list copied from Excel, a mail or pictures of it is pasted into
 * the AI box and AI fills the form; the form is changed in place, saved, and printed as the paper form.
 */
export default function RequestPage(props: {
  today: string; onError: (e: unknown) => void;
  /** Already loaded (tests); loaded here when left out. */
  initial?: WorkingRequest[];
  /** The request open at first (else a new one). */
  openId?: string;
  /** Drawing numbers already checked (tests); else asked of the drawing DB as they are typed. */
  checked?: DrawingCheck[];
}) {
  const tr = useTr();
  const [list, setList] = useState<WorkingRequest[] | null>(props.initial ?? null);
  const first = props.initial?.find((r) => r.id === props.openId);
  const [id, setId] = useState<string | null>(first?.id ?? null);
  const [form, setForm] = useState<RequestFields>(() => (first ? readRequest(first, props.today) : blankRequest(props.today)));
  const [dirty, setDirty] = useState(false);
  const [saving, setSaving] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [problem, setProblem] = useState("");
  // AI box: what is pasted, the pictures, the run going on and what AI said last
  const [text, setText] = useState("");
  const [pictures, setPictures] = useState<File[]>([]);
  const [busy, setBusy] = useState<AbortController | null>(null);
  const [said, setSaid] = useState<{ note: string; model: string } | null>(null);
  // the model typed once above the rows, the drawing numbers the drawing DB answered for (by drawingKey), what went
  // wrong reaching the DB, and the row whose drawing number is being picked from the suggestions
  const [model, setModel] = useState("");
  const [checks, setChecks] = useState<Map<string, DrawingCheck>>(() => new Map((props.checked ?? []).map((c) => [drawingKey(c.drawingNo), c])));
  const [dbProblem, setDbProblem] = useState("");
  const [picking, setPicking] = useState<number | null>(null);

  // the drawing numbers not checked yet are asked of the drawing DB a moment after typing stops
  const asking = uncheckedNos(form.items, checks).join("\n");
  useEffect(() => {
    if (!asking || props.checked) return;
    const ctl = new AbortController();
    const t = window.setTimeout(() => {
      drawingApi.check(asking.split("\n"), ctl.signal).then((r) => {
        if (!r.ok) { setDbProblem(r.error || "The drawing DB could not be opened."); return; }
        setDbProblem("");
        setChecks((cur) => { const n = new Map(cur); for (const c of r.results) n.set(drawingKey(c.drawingNo), c); return n; });
      }, (e) => { if (!ctl.signal.aborted) setDbProblem(e instanceof Error ? e.message : String(e)); });
    }, 600);
    return () => { window.clearTimeout(t); ctl.abort(); };
  }, [asking]);

  useEffect(() => {
    if (props.initial) return;
    let live = true;
    requestApi.list().then((l) => { if (live) setList(l); }, (e) => { if (live) { setList([]); props.onError(e); } });
    return () => { live = false; };
  }, []);

  const change = (f: Partial<RequestFields>) => { setForm((cur) => ({ ...cur, ...f })); setDirty(true); setProblem(""); };
  const changeItem = (n: number, f: Partial<RequestItem>) => change({ items: form.items.map((i, k) => (k === n ? { ...i, ...f } : i)) });
  const addItem = () => change({ items: [...form.items, blankItem(form.items.length + 1, model.trim() || commonModel(form.items) || form.partName)] });
  /** The model above the rows put into the rows without one (or every row); the part name too when it has none. */
  const putModel = (every: boolean) => {
    if (!model.trim()) return;
    change({ items: withModel(form.items, model, every), ...(form.partName.trim() ? {} : { partName: model.trim() }) });
  };
  const removeItem = (n: number) => change({ items: form.items.filter((_, k) => k !== n).map((i, k) => ({ ...i, no: k + 1 })) });

  const leave = () => !dirty || window.confirm(tr("Discard the unsaved changes?"));
  const open = (r: WorkingRequest | null) => {
    if (!leave()) return;
    setId(r?.id ?? null);
    setForm(r ? readRequest(r, props.today) : blankRequest(props.today));
    setDirty(false); setProblem(""); setSaid(null);
  };

  const save = async () => {
    const p = requestProblem(form);
    if (p) { setProblem(p); return; }
    setSaving(true); setProblem("");
    try {
      const saved = id ? await requestApi.save(id, form) : await requestApi.create(form);
      setList((cur) => [saved, ...(cur ?? []).filter((r) => r.id !== saved.id)]);
      setId(saved.id); setForm(readRequest(saved, props.today)); setDirty(false);
    } catch (e) { setProblem(e instanceof Error ? e.message : String(e)); } finally { setSaving(false); }
  };
  /** The form as it is on the screen (saved or not) downloaded as an .xlsx. */
  const exportExcel = async () => {
    setExporting(true); setProblem("");
    try { downloadBlob(await requestApi.exportXlsx(form), requestFileName(form)); }
    catch (e) { setProblem(e instanceof Error ? e.message : String(e)); } finally { setExporting(false); }
  };
  const remove = async () => {
    if (!id || !window.confirm(tr("Delete this request?"))) return;
    try {
      await requestApi.remove(id);
      setList((cur) => (cur ?? []).filter((r) => r.id !== id));
      setId(null); setForm(blankRequest(props.today)); setDirty(false);
    } catch (e) { props.onError(e); }
  };

  /** Pictures pasted or dropped are kept to send; a text paste goes on into the box. */
  const addPictures = (files: File[]) => {
    const r = aiPictures(files, pictures.length, MAX_REQUEST_PICTURE_BYTES);
    if (r.problem) setProblem(r.problem);
    else if (r.files.length) { setProblem(""); setPictures((cur) => [...cur, ...r.files]); }
    return r.files.length > 0 || !!r.problem;
  };
  const onPaste = (ev: ClipboardEvent) => { if (!busy && addPictures(Array.from(ev.clipboardData.files))) ev.preventDefault(); };
  const onDrop = (ev: DragEvent) => { if (ev.dataTransfer.files.length) { ev.preventDefault(); if (!busy) addPictures(Array.from(ev.dataTransfer.files)); } };

  const fill = async () => {
    if ((!text.trim() && !pictures.length) || busy) return;
    const ctl = new AbortController();
    setBusy(ctl); setProblem(""); setSaid(null);
    try {
      const images = await Promise.all(pictures.map(async (f, i) => ({ name: f.name || `pasted-${i + 1}.png`, dataUrl: await dataUrl(f) })));
      const a = await requestApi.ai({ text, today: props.today, kind: form.kind, current: form, images }, ctl.signal);
      const next = readRequest(a.request, props.today);
      if (!next.partName) next.partName = commonModel(next.items);
      setForm(requestFields(next));
      setDirty(true);
      setSaid({ note: a.note ?? "", model: a.model ?? "" });
      setText(""); setPictures([]);
    } catch (e) {
      if (!ctl.signal.aborted) setProblem(e instanceof Error ? e.message : String(e));
    } finally { setBusy(null); }
  };

  const sum = totals(form.items);
  const states = form.items.map((i) => checkOf(i.drawingNo, checks)?.status);
  const missing = states.filter((x) => x === "missing").length;
  const revised = states.filter((x) => x === "revision").length;
  const noDrawing = form.items.filter((i) => i.name.trim() && !i.drawingNo.trim()).length;
  const field = (k: keyof RequestFields, cls = "", ph = "") => (
    <input className={"rq-in " + cls} value={form[k] as string} placeholder={ph} aria-label={k} onChange={(e) => change({ [k]: e.target.value })} />
  );
  const date = (k: "requestDate" | "finishDate1" | "finishDate2") => (
    <>
      <input type="date" className="rq-in rq-date rq-noprint" value={form[k]} aria-label={k} onChange={(e) => change({ [k]: e.target.value })} />
      <span className="rq-print">{dotDate(form[k])}</span>
    </>
  );
  const label = (t: ReactNode) => <th colSpan={2} className="rq-l">{t}</th>;

  return (
    <div className="rq-page">
      {/* the margins of the form when it is printed (only while this page is open); the paper and its direction are the
          print dialog's, a portrait page getting the form smaller (request.css) */}
      <style>{"@page { margin: 8mm; }"}</style>
      <aside className="card rq-list rq-noprint">
        <div className="rq-list-h">
          <h2>{tr("Requests")}</h2>
          <button className="btn sm" onClick={() => open(null)}>+ {tr("New")}</button>
        </div>
        {list === null && <div className="muted small">Loading…</div>}
        {list?.length === 0 && <div className="muted small">{tr("No request yet.")}</div>}
        {list?.map((r) => (
          <button key={r.id} className={"rq-li" + (r.id === id ? " on" : "")} onClick={() => r.id !== id && open(r)}>
            <b>{requestTitle(r)}</b>
            <span className="muted small">{[r.kind, dotDate(r.requestDate), `${r.items.length} items`, r.updatedBy].filter(Boolean).join(" · ")}</span>
          </button>
        ))}
      </aside>

      <div className="rq-main">
        <div className="card rq-ai rq-noprint">
          <div className="rq-ai-h">
            <b>✦ {tr("AI input")}</b>
            <select className="inp sm" value={form.kind} aria-label="Kind" onChange={(e) => change({ kind: e.target.value })}>
              {REQUEST_KINDS.map((k) => <option key={k}>{k}</option>)}
            </select>
            <span className="muted small">
              {tr("Paste the list (cells copied from Excel, a mail) or pictures of it (Ctrl+V or drop; up to 6). AI fills the form below. Ctrl+Enter sends.")}
            </span>
          </div>
          <textarea className="inp wk-ta wk-ai-text" rows={4} value={text} disabled={!!busy} aria-label="List to make the request of"
            placeholder={"e.g.\nSM-A386TOP\tMAIN 1 JIG-Y\t40\t0\nSM-A386TOP\tF-PCB SOLDER CAP JIG\t2\t0\nUsing place C2, finish 11/20"}
            onPaste={onPaste} onDrop={onDrop} onDragOver={(e) => e.preventDefault()}
            onChange={(e) => setText(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { e.preventDefault(); void fill(); } }} />
          {pictures.length > 0 && (
            <div className="wk-ai-sent">
              <span className="muted small">{pictures.length} picture{pictures.length > 1 ? "s" : ""} to send</span>
              <Thumbs files={pictures} onRemove={busy ? undefined : (i) => setPictures((cur) => cur.filter((_, n) => n !== i))} />
            </div>
          )}
          <div className="wk-ai-bar">
            <button className="btn sm primary" disabled={(!text.trim() && !pictures.length) || !!busy} onClick={() => void fill()}>
              {busy ? tr("AI is reading…") : tr("Fill the form with AI")}
            </button>
            {busy && <button className="btn sm" onClick={() => busy.abort()}>{tr("Cancel")}</button>}
            {said?.model && <span className="muted small">{said.model}</span>}
          </div>
          {said?.note && <div className="wk-ai-note rq-said">{said.note}</div>}
        </div>

        <div className="rq-bar rq-noprint">
          <button className="btn sm primary" disabled={saving || (!dirty && !!id)} onClick={() => void save()}>{saving ? tr("Saving…") : tr("Save")}</button>
          <button className="btn sm" onClick={() => window.print()}>{tr("Print")}</button>
          <button className="btn sm" disabled={exporting} onClick={() => void exportExcel()} title="Download the form as Excel (.xlsx)">
            {exporting ? tr("Exporting…") : tr("Export Excel")}
          </button>
          {id && <button className="btn sm danger" onClick={() => void remove()}>{tr("Delete")}</button>}
          {dirty && <span className="pi-dirty small">{tr("Not saved")}</span>}
          {problem && <span className="pi-dirty small" role="alert">{problem}</span>}
        </div>

        <div className="rq-sheet">
          <div className="rq-tools rq-noprint">
            <label className="rq-model">
              <b>{tr("Model")}</b>
              <input className="inp sm" value={model} placeholder="e.g. MSU-L15S10-20" aria-label="Model of the rows"
                onChange={(e) => setModel(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter") putModel(false); }} />
            </label>
            <button className="btn sm" disabled={!model.trim() || !form.items.length} onClick={() => putModel(false)}
              title="Put the model into the rows that have none">{tr("Fill empty rows")}</button>
            <button className="btn sm" disabled={!model.trim() || !form.items.length} onClick={() => putModel(true)}
              title="Put the model into every row">{tr("All rows")}</button>
            <span className="rq-dwg-sum">
              {dbProblem ? <span className="pi-dirty small" role="alert">Drawing DB: {dbProblem}</span> : (
                <>
                  {missing > 0 && <span className="rq-tag miss">{missing} {tr("not in the drawing DB")}</span>}
                  {revised > 0 && <span className="rq-tag rev">{revised} {tr("other revision")}</span>}
                  {noDrawing > 0 && <span className="rq-tag none">{noDrawing} {tr("without a drawing no (click the magnifier to pick)")}</span>}
                </>
              )}
            </span>
          </div>
          <table className="rq-form">
            <colgroup>{[110, 223, 182, 182, 191, 87, 85, 87, 173, 169, 170, 169, 170].map((w, i) => <col key={i} style={{ width: `${w / 19.98}%` }} />)}</colgroup>
            <tbody>
              <tr>
                <th rowSpan={2} className="rq-l">Request<br />depart</th>
                <th className="rq-s">Request</th><th className="rq-s">Phone</th><th className="rq-s">Checker</th><th className="rq-s">Confirm</th>
                <td colSpan={4} rowSpan={2} className="rq-title">PROCESSING REQUEST</td>
                <th rowSpan={2} className="rq-l">Manufacture<br />Department</th>
                <th className="rq-s">Design</th><th className="rq-s">Approval</th><th className="rq-s">Confirm</th>
              </tr>
              <tr className="rq-sign">
                <td>{field("requester")}</td><td>{field("phone")}</td><td>{field("checker")}</td><td>{field("confirm")}</td>
                <td /><td /><td />
              </tr>
              <tr>
                {label("Processing method")}
                <td colSpan={7} className="rq-methods">
                  {METHODS.map(([m, t]) => (
                    <label key={m} className="rq-method">
                      <input type="checkbox" checked={form.method === m} onChange={() => change({ method: m })} />
                      <span className="rq-box">{form.method === m ? "✔" : ""}</span><span className="rq-mname">{t}</span>
                    </label>
                  ))}
                </td>
                <th colSpan={2} className="rq-l">Department name</th>
                <td colSpan={2}>{field("departmentName", "rq-b")}</td>
              </tr>
              <tr>
                {label("*Equipment Name")}
                <td colSpan={7}>{field("equipmentName")}</td>
                <th colSpan={2} rowSpan={2} className="rq-l">*Quantity</th>
                <th className="rq-s">1st time</th><th className="rq-s">2nd time</th>
              </tr>
              <tr>
                {label("*Part Name")}
                <td colSpan={7}>{field("partName", "rq-part")}</td>
                <td>{field("qty1")}</td><td>{field("qty2")}</td>
              </tr>
              <tr>
                <th colSpan={2} rowSpan={2} className="rq-l">*Item Application</th>
                <td colSpan={7} rowSpan={2}>{field("itemApplication", "rq-big")}</td>
                <th colSpan={2} className="rq-l">*Request date</th>
                <td colSpan={2} className="rq-b">{date("requestDate")}</td>
              </tr>
              <tr>
                <th colSpan={2} className="rq-l">*Purpose</th>
                <td colSpan={2}>{field("purpose", "rq-b")}</td>
              </tr>
              <tr>
                {label("*Using place")}
                <td colSpan={7}>{field("usingPlace", "rq-big")}</td>
                <th colSpan={2} rowSpan={2} className="rq-l rq-u">*Finish date 2nd time</th>
                <td colSpan={2} rowSpan={2} className="rq-b">{date("finishDate2")}</td>
              </tr>
              <tr>
                {label(<span className="rq-u">*Finish date<br />1st time</span>)}
                <td colSpan={7} className="rq-b">{date("finishDate1")}</td>
              </tr>
              <tr>
                {label("*Others")}
                <td colSpan={7}>{field("others", "", "Part(   ) drawing")}</td>
                <th colSpan={2} className="rq-l">*Dep&apos;t code</th>
                <td colSpan={2}>{field("deptCode", "rq-code", "auto")}</td>
              </tr>
              <tr className="rq-head">
                <th>No.</th><th>Model</th><th colSpan={3}>Name</th><th>Line<br />Spec</th><th>Stock</th><th>Order</th>
                <th colSpan={2}>Picture/Drawing no</th><th>PICTURE</th><th colSpan={2}>Note</th>
              </tr>
              {form.items.map((i, n) => (
                <tr key={n} className="rq-item">
                  <td className="rq-no">{n + 1}<button className="rq-x rq-noprint" aria-label={`Remove row ${n + 1}`} onClick={() => removeItem(n)}>×</button></td>
                  <td><input className="rq-in" value={i.model} aria-label={`Model ${n + 1}`} onChange={(e) => changeItem(n, { model: e.target.value })} /></td>
                  <td colSpan={3}><input className="rq-in" value={i.name} aria-label={`Name ${n + 1}`} onChange={(e) => changeItem(n, { name: e.target.value })} /></td>
                  {(["lineSpec", "stock"] as const).map((k) => (
                    <td key={k}><input className="rq-in rq-num" type="number" min={0} value={i[k] || 0} aria-label={`${k} ${n + 1}`}
                      onChange={(e) => changeItem(n, { [k]: Math.max(0, Number(e.target.value) || 0) })} /></td>
                  ))}
                  <td>
                    <input className="rq-in rq-num rq-noprint" type="number" min={0} value={i.order || ""} placeholder={String(orderOf(i))} aria-label={`order ${n + 1}`}
                      onChange={(e) => changeItem(n, { order: Math.max(0, Number(e.target.value) || 0) })} />
                    <span className="rq-print">{orderOf(i)}</span>
                  </td>
                  <td colSpan={2} className={"rq-dwg" + dwgClass(checkOf(i.drawingNo, checks))} title={dwgTitle(checkOf(i.drawingNo, checks))}>
                    <input className="rq-in rq-noprint" value={i.drawingNo} aria-label={`Drawing no ${n + 1}`} onChange={(e) => changeItem(n, { drawingNo: e.target.value })} />
                    <span className="rq-print">{i.drawingNo}</span>
                    {(!i.drawingNo.trim() || checkOf(i.drawingNo, checks)?.status === "missing") && i.name.trim() ? (
                      <button className="rq-find rq-noprint" aria-label={`Find the drawing no of row ${n + 1}`} title="Pick a drawing no from the drawing DB"
                        onClick={() => setPicking(n)}>&#128269;</button>
                    ) : otherVersions(i.drawingNo, checkOf(i.drawingNo, checks)).length > 0 && (
                      <button className="rq-find rq-noprint" aria-label={`Other versions of the drawing no of row ${n + 1}`}
                        title={`Other versions in the drawing DB: ${otherVersions(i.drawingNo, checkOf(i.drawingNo, checks)).map((v) => v.drawingNo).join(", ")}`}
                        onClick={() => setPicking(n)}>&#128269;</button>
                    )}
                  </td>
                  <td><input className="rq-in" value={i.picture} aria-label={`Picture ${n + 1}`} onChange={(e) => changeItem(n, { picture: e.target.value })} /></td>
                  <td colSpan={2}>
                    <input className="rq-in rq-note rq-noprint" value={i.note} title={i.note} aria-label={`Note ${n + 1}`} onChange={(e) => changeItem(n, { note: e.target.value })} />
                    <span className="rq-print rq-print-wrap">{i.note}</span>
                  </td>
                </tr>
              ))}
              {form.items.length > 1 && (
                <tr className="rq-total">
                  <td colSpan={5}>Total</td><td>{sum.lineSpec}</td><td>{sum.stock}</td><td>{sum.order}</td><td colSpan={5} />
                </tr>
              )}
            </tbody>
          </table>
          <button className="btn sm rq-add rq-noprint" onClick={addItem}>+ {tr("Add row")}</button>
          {picking !== null && form.items[picking] && (
            <DrawingPicker item={form.items[picking]} model={form.items[picking].model || model}
              versions={checkOf(form.items[picking].drawingNo, checks)?.versions}
              onPick={(m) => {
                // a version picked keeps the versions of its base, so the others can still be found from it
                const versions = checkOf(form.items[picking].drawingNo, checks)?.versions?.some((v) => drawingKey(v.drawingNo) === drawingKey(m.drawingNo))
                  ? checkOf(form.items[picking].drawingNo, checks)?.versions : undefined;
                changeItem(picking, { drawingNo: m.drawingNo });
                setChecks((cur) => new Map(cur).set(drawingKey(m.drawingNo), { drawingNo: m.drawingNo, status: "found", match: m, versions }));
                setPicking(null);
              }}
              onClose={() => setPicking(null)} />
          )}
        </div>
      </div>
    </div>
  );
}

/** How the drawing-number cell looks: red when the DB has no such drawing, amber when only another revision. */
const dwgClass = (c: DrawingCheck | undefined) => (c?.status === "missing" ? " miss" : c?.status === "revision" ? " rev" : c?.status === "found" ? " ok" : "");
const dwgTitle = (c: DrawingCheck | undefined) =>
  !c ? undefined : c.status === "missing" ? "Not in the JIG/MC drawing DB"
    : c.status === "revision" ? `Only another revision is in the drawing DB: ${matchText(c.match)}` : matchText(c.match);
