import { useEffect, useState } from "react";
import { bmesApi, drawingApi, drawingsByName, drawingsOf, matchMaterial, qty, shiftCount, subGroupTree, suggestMaterials, visibleBom } from "./bmes";
import type { BmesMaterial, BomResult, PdmRow, SubGroupNode } from "./bmes";
import { useTr } from "./i18n";
import ManualGroup from "./ManualGroup";
import { MANUAL, emptyManual, isEmptyManual, keepSubTables, shownTable, type ManualGroup as ManualGroupData } from "./manual";
import TempGroupEdit from "./TempGroupEdit";
import { hasTempGroup, subPaths, type TempGroup as TempGroupData } from "./tempGroup";

/** Links to the drawings of a material (opened in a new tab), or a dash. */
function Drawings(props: { list: PdmRow[] }) {
  if (!props.list.length) return <span className="muted">—</span>;
  return (
    <>
      {props.list.map((d) => (
        <a key={d.pdmno + "/" + d.pdmvr} className="wk-dwg" href={drawingApi.url(d)} target="_blank" rel="noreferrer"
          title={`Open drawing ${d.pdmno} rev ${d.pdmvr || "0"}${d.pdtnoTx ? ` (${d.pdtnoTx})` : ""}`}>
          {d.pdmno}{d.pdmvr && d.pdmvr !== "0" ? ` r${d.pdmvr}` : ""}
        </a>
      ))}
    </>
  );
}

/** The Create Date or Mod. Date of each drawing of a material (date only, the time in its tip), a line each as its links are. */
export function DrawingDates(props: { list: PdmRow[]; field: "erdat" | "ardat" }) {
  if (!props.list.length) return <span className="muted">—</span>;
  return (
    <>
      {props.list.map((d) => {
        const v = d[props.field] ?? "";
        return <div key={d.pdmno + "/" + d.pdmvr} className="wk-dwg-date" title={v}>{v.slice(0, 10) || "—"}</div>;
      })}
    </>
  );
}

/** BOM of one product with the material names and their drawings: from the local cache, or loaded from BMES on request. */
function BomTable(props: { product: string; onError: (e: unknown) => void }) {
  const { product } = props;
  const [bom, setBom] = useState<BomResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [closed, setClosed] = useState<Set<number>>(new Set());
  const [drawings, setDrawings] = useState<Map<string, PdmRow[]>>(new Map());
  const [drawingError, setDrawingError] = useState("");

  useEffect(() => {
    let live = true;
    setBom(null);
    setClosed(new Set());
    bmesApi.bom(product).then((b) => { if (live) setBom(b); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [product]);

  // drawings of the product and of every material in its BOM
  const names = bom ? [product, ...bom.rows.map((r) => r.materialName)].join("\n") : "";
  useEffect(() => {
    if (!names) return;
    let live = true;
    setDrawingError("");
    drawingApi.find([...new Set(names.split("\n"))])
      .then((rows) => { if (live) setDrawings(drawingsByName(rows)); })
      .catch((e: Error) => { if (live) setDrawingError(e.message); });
    return () => { live = false; };
  }, [names]);

  const load = async () => {
    setBusy(true);
    try { setBom(await bmesApi.fetchBom(product)); setClosed(new Set()); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const toggle = (i: number) => setClosed((s) => { const n = new Set(s); if (n.has(i)) n.delete(i); else n.add(i); return n; });

  return (
    <div className="wk-bom">
      <div className="wk-bom-h">
        <b>{product}</b>
        <span className="wk-bom-own">Drawing <Drawings list={drawingsOf(drawings, product)} /></span>
        {bom === null ? <span className="muted small">Loading…</span>
          : bom.rows.length ? (
            <span className="muted small">
              {bom.rows.length} rows · BOM date {bom.workDate || "—"} · {bom.source === "server" ? "loaded from BMES" : "from the local BOM cache"}
              {bom.fetchedAt && ` ${bom.fetchedAt.slice(0, 16).replace("T", " ")}`}
            </span>
          ) : <span className="muted small">No BOM saved for this product yet.</span>}
        <button className="btn sm" disabled={busy || bom === null} onClick={load}
          title="Read the BOM of this product from BMES (read only) and keep it in the local cache">
          {busy ? "Loading from BMES…" : bom?.rows.length ? "Reload from BMES" : "Load from BMES"}
        </button>
      </div>
      {bom?.error && <div className="pi-dirty">{bom.error}</div>}
      {drawingError && <div className="muted small">Drawings could not be looked up: {drawingError}</div>}
      {!!bom?.rows.length && (
        <div className="pi-wrap">
          <table className="wk-it wk-bom-t">
            <thead>
              <tr>
                <th>Level</th><th className="l">Material name</th><th>Material code</th><th>Qty</th><th>Unit</th><th className="l">Drawing</th>
                <th title="When the drawing was registered in BMES (ERDAT)">Create Date</th>
                <th title="When the drawing was last changed in BMES (ARDAT)">Mod. Date</th>
              </tr>
            </thead>
            <tbody>
              {visibleBom(bom.rows, closed).map(({ row, index }) => (
                <tr key={index} className={"lv" + Math.min(row.level, 6)}>
                  <td className="n">{row.level}</td>
                  <td className="l" style={{ paddingLeft: 8 + (row.level - 1) * 18 }}>
                    {row.hasChildren
                      ? <button className="wk-bom-tog" onClick={() => toggle(index)} aria-expanded={!closed.has(index)}>{closed.has(index) ? "▸" : "▾"}</button>
                      : <span className="wk-bom-leaf" />}
                    {row.materialName}
                  </td>
                  <td className="wk-bom-code">{row.materialCode}</td>
                  <td className="n">{qty(row.usageQty)}</td>
                  <td>{row.usageUnit}</td>
                  <td className="l wk-bom-dwg"><Drawings list={drawingsOf(drawings, row.materialName)} /></td>
                  <td><DrawingDates list={drawingsOf(drawings, row.materialName)} field="erdat" /></td>
                  <td><DrawingDates list={drawingsOf(drawings, row.materialName)} field="ardat" /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/** Shift groups in columns. */
const Shifts = (props: { list: string[] }) => (
  <div className="wk-bg-shifts">
    {props.list.map((ls) => <div key={ls} className="wk-bg-shift"><span className="wk-bg-sbadge">Shift Group</span>{ls}</div>)}
  </div>
);

/** One sub-group as BMES > Make Model Group shows it: its name, how many line shifts, the shift groups, the sub-groups under it. */
function SubGroupCard(props: { node: SubGroupNode }) {
  const n = props.node;
  return (
    <div className="wk-bg-sub">
      <div className="wk-bg-row">
        <span className="wk-bg-badge">Sub Group {n.depth}</span><b>{n.name}</b>
        <span className="muted small wk-bg-count">{shiftCount(n)} shifts</span>
      </div>
      {n.lineShifts.length > 0 && <Shifts list={n.lineShifts} />}
      {n.children.length > 0 && <div className="wk-bg-subs">{n.children.map((c) => <SubGroupCard key={c.name} node={c} />)}</div>}
    </div>
  );
}

/** The model of BMES > Make Model Group the Working model is linked to, laid out as that page shows it. */
function LinkedGroup(props: { m: BmesMaterial }) {
  const { m } = props;
  const tree = subGroupTree(m.subGroups);
  return (
    <div className="wk-bg-model">
      <div className="wk-bg-row">
        <span className="wk-bg-badge model">Model</span><b>{m.material}</b>
        <span className="wk-chip on">{m.group}</span>
        <span className="muted small wk-bg-count">{tree.nodes.length} sub groups · {m.subGroups.reduce((s, x) => s + x.lineShifts.length, 0)} shifts</span>
      </div>
      {tree.direct.length > 0 && <Shifts list={tree.direct} />}
      <div className="wk-bg-subs">{tree.nodes.map((n) => <SubGroupCard key={n.name} node={n} />)}</div>
      {!m.subGroups.length && <div className="muted small">No line shift in this model.</div>}
    </div>
  );
}

/** The models of BMES > Make Model Group by group, as that page lists them, to link one to the Working model. */
function GroupPicker(props: {
  modelName: string; materials: BmesMaterial[]; current: string;
  onLink: (material: string) => void; onManual: () => void; onCancel?: () => void;
}) {
  const groups = [...new Set(props.materials.map((x) => x.group))];
  const [group, setGroup] = useState("");
  const [query, setQuery] = useState("");
  const q = query.trim().toLowerCase();
  const suggested = suggestMaterials(props.modelName, props.materials);
  const list = props.materials.filter((x) => (!group || x.group === group) && (!q || x.material.toLowerCase().includes(q)));
  const row = (x: BmesMaterial) => (
    <div key={x.group + "/" + x.material} className={"wk-bg-pick" + (x.material === props.current ? " on" : "")}>
      <span className="wk-bg-badge model">Model</span><b>{x.material}</b>
      <span className="muted small">{subGroupTree(x.subGroups).nodes.length} sub groups · {x.subGroups.reduce((s, y) => s + y.lineShifts.length, 0)} shifts</span>
      <button className="btn xs primary" onClick={() => props.onLink(x.material)}>Link</button>
    </div>
  );
  return (
    <div className="wk-bg-picker">
      {suggested.length > 0 && !q && !group && (
        <div className="wk-bg-group sugg">
          <div className="wk-bg-gname">Looks like {props.modelName}</div>
          {suggested.map(row)}
        </div>
      )}
      <div className="wk-bg-filter">
        <span className="muted small">PRODUCT</span>
        <button className={"wk-chip" + (!group ? " on" : "")} onClick={() => setGroup("")}>All</button>
        {groups.map((g) => <button key={g} className={"wk-chip" + (g === group ? " on" : "")} onClick={() => setGroup(g)}>{g}</button>)}
        <input className="inp sm" placeholder="Search model…" value={query} aria-label="Search BMES models" onChange={(e) => setQuery(e.target.value)} />
      </div>
      <div className="wk-bg-groups">
        {groups.filter((g) => list.some((x) => x.group === g)).map((g) => (
          <div key={g} className="wk-bg-group">
            <div className="wk-bg-gname">{g}</div>
            {list.filter((x) => x.group === g).map(row)}
          </div>
        ))}
      </div>
      {!list.length && <div className="muted">No BMES model matches.</div>}
      <div className="wk-bg-foot">
        {props.onCancel && <button className="btn sm" onClick={props.onCancel}>Cancel</button>}
        <button className="wk-link" onClick={props.onManual}>Not in BMES yet? Make a temporary model group here (sub groups, shift groups)</button>
      </div>
    </div>
  );
}

/**
 * BMES model group of the model (BMES > Make Model Group), shown inside Model information as that page shows it: found
 * by the model name or picked from the Make Model Group list and saved with the model information, with its sub-groups
 * and shift groups (the BOM of each product is on the BOM / Drawing tab, BomPanel). A model BMES does not have yet gets a temporary group made here by hand
 * (Model › Sub Group › Shift Group) with its processes, NG items and materials.
 */
/**
 * Temporary model group of a model BMES does not have yet: its Model › Sub Group › Shift Group tree. The processes, NG items
 * and materials are on the Process tab (ProcessTables).
 */
export function TempModelGroup(props: {
  modelName: string;
  temp: TempGroupData; onTemp: (g: TempGroupData) => void;
  subManual: Record<string, ManualGroupData>; onSubManual: (t: Record<string, ManualGroupData>) => void;
  onError: (e: unknown) => void;
}) {
  // a removed sub-group takes its process table with it
  const onTemp = (t: TempGroupData) => {
    props.onTemp(t);
    const kept = keepSubTables(props.subManual, subPaths(t).map((s) => s.id));
    if (kept !== props.subManual) props.onSubManual(kept);
  };
  return (
    <>
      <TempGroupEdit modelName={props.modelName} value={props.temp} onChange={onTemp} onError={props.onError} />
      <div className="muted small wk-tg-note">Processes, NG items and materials are on the Process tab.</div>
    </>
  );
}

/**
 * Process tab: the processes, NG items and materials of the model, or of the sub-group picked (chips) when the temporary
 * model group has sub-groups (the model's own table is offered while there is no sub-group or it already has something).
 */
export function ProcessTables(props: {
  modelName: string;
  manual: ManualGroupData; onManual: (g: ManualGroupData) => void;
  temp: TempGroupData;
  subManual: Record<string, ManualGroupData>; onSubManual: (t: Record<string, ManualGroupData>) => void;
  onError: (e: unknown) => void;
  /** Table picked first (tests); else the first offered. */
  initial?: string;
}) {
  const tr = useTr();
  // process table picked ("" = the model's, else a sub-group id; null = not picked yet)
  const [table, setTable] = useState<string | null>(props.initial ?? null);
  const subs = subPaths(props.temp);
  const tbl = shownTable(table, subs.map((s) => s.id), props.manual);
  const sub = subs.find((s) => s.id === tbl.at);
  return (
    <>
      {subs.length > 0 && (
        <div className="wk-tg-tables">
          <span className="muted small">{tr("Processes and materials of")}</span>
          {tbl.model && <button className={"wk-chip" + (tbl.at === "" ? " on" : "")} onClick={() => setTable("")}>{tr("Model")} {props.modelName}</button>}
          {subs.map((s) => (
            <button key={s.id} className={"wk-chip" + (s.id === tbl.at ? " on" : "")} onClick={() => setTable(s.id)}>{s.path}</button>
          ))}
        </div>
      )}
      {sub
        ? <ManualGroup key={sub.id} modelName={props.modelName} product={sub.path.split(" › ").pop()} value={props.subManual[sub.id] ?? emptyManual()}
            onChange={(g) => props.onSubManual({ ...props.subManual, [sub.id]: g })} onError={props.onError} />
        : <ManualGroup key="" modelName={props.modelName} value={props.manual} onChange={props.onManual} onError={props.onError} />}
    </>
  );
}

export default function BmesPanel(props: {
  modelName: string;
  /** Material saved in Model Info ("" = found by name). */
  saved: string;
  onPick: (material: string) => void;
  /** Process table of the model (read only here: whether the model has a hand-made group; it is edited on the Process tab). */
  manual: ManualGroupData;
  /** Model › Sub Group › Shift Group made by hand for a model BMES does not have yet. */
  temp: TempGroupData;
  onTemp: (g: TempGroupData) => void;
  /** Processes, NG items and materials of each sub-group of the temporary group, by sub-group id. */
  subManual: Record<string, ManualGroupData>;
  onSubManual: (t: Record<string, ManualGroupData>) => void;
  onError: (e: unknown) => void;
}) {
  const [materials, setMaterials] = useState<BmesMaterial[] | null>(null);
  const [picking, setPicking] = useState(false);

  useEffect(() => {
    let live = true;
    bmesApi.materials().then((m) => { if (live) setMaterials(m); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const isManual = props.saved === MANUAL;
  const found = materials && !isManual ? matchMaterial(props.modelName, materials, props.saved) : null;
  const m = found?.material ?? null;
  const hasManual = !isEmptyManual(props.manual) || hasTempGroup(props.temp) || Object.values(props.subManual).some((t) => !isEmptyManual(t));
  const link = (material: string) => { props.onPick(material); setPicking(false); };

  let body;
  if (materials === null) body = <div className="muted">Loading the BMES model groups…</div>;
  else if (picking || (!m && !isManual && !hasManual)) {
    body = <GroupPicker modelName={props.modelName} materials={materials} current={m?.material ?? ""} onLink={link}
      onManual={() => link(MANUAL)} onCancel={picking ? () => setPicking(false) : undefined} />;
  } else if (!m) {
    body = <TempModelGroup modelName={props.modelName} temp={props.temp} onTemp={props.onTemp}
      subManual={props.subManual} onSubManual={props.onSubManual} onError={props.onError} />;
  } else {
    body = <LinkedGroup m={m} />;
  }

  return (
    <div className="wk-bmes">
      <div className="wk-bmes-h">
        <h3>BMES model group</h3>
        {materials !== null && !picking && (m || isManual || hasManual) && (
          <span className="wk-bmes-pick">
            {m ? <>{m.group} › <b>{m.material}</b> · {found?.by === "saved" ? "linked in Model Info" : "same name"}</> : null}
            <button className="btn sm" onClick={() => setPicking(true)}>Change link</button>
          </span>
        )}
      </div>
      {body}
    </div>
  );
}

/**
 * BOM / Drawing tab of a model: the BOM of each product of the BMES model group linked in Model Info (or of the same
 * name), with the drawings of its materials and when each was registered and last changed.
 */
export function BomPanel(props: {
  modelName: string;
  /** Material saved in Model Info ("" = found by name). */
  saved: string;
  onError: (e: unknown) => void;
  /** BMES model groups already read (tests); else read here. */
  initial?: BmesMaterial[];
}) {
  const [materials, setMaterials] = useState<BmesMaterial[] | null>(props.initial ?? null);
  const [product, setProduct] = useState("");

  useEffect(() => {
    if (props.initial) return;
    let live = true;
    bmesApi.materials().then((m) => { if (live) setMaterials(m); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const m = materials && props.saved !== MANUAL ? matchMaterial(props.modelName, materials, props.saved)?.material ?? null : null;
  const shown = m && m.products.includes(product) ? product : m?.products[0] ?? "";
  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>BOM / Drawing</h2>
        {m && <span className="muted small">{m.group} › <b>{m.material}</b></span>}
      </div>
      <div className="pi-sec-b">
        {materials === null ? <div className="muted">Loading the BMES model groups…</div>
          : !m ? <div className="muted">No BMES model group is linked to this model. Link it in Model Info to see its BOM and drawings.</div>
            : (
              <>
                <div className="wk-bmes-bomh">
                  <h3>BOM</h3>
                  {m.products.map((p) => (
                    <button key={p} className={"wk-chip" + (p === shown ? " on" : "")} onClick={() => setProduct(p)}>{p}</button>
                  ))}
                </div>
                {shown ? <BomTable key={shown} product={shown} onError={props.onError} /> : <div className="muted">No product in this model.</div>}
              </>
            )}
      </div>
    </section>
  );
}
