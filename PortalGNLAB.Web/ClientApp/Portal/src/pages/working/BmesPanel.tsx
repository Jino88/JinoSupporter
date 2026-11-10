import { useEffect, useState } from "react";
import { bmesApi, drawingApi, drawingsByName, drawingsOf, matchMaterial, qty, shiftCount, subGroupTree, suggestMaterials, visibleBom } from "./bmes";
import type { BmesMaterial, BomResult, PdmRow, SubGroupNode } from "./bmes";
import ManualGroup from "./ManualGroup";
import { MANUAL, type ManualGroup as ManualGroupData } from "./manual";

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
              <tr><th>Level</th><th className="l">Material name</th><th>Material code</th><th>Qty</th><th>Unit</th><th className="l">Drawing</th></tr>
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
                  <td className="l"><Drawings list={drawingsOf(drawings, row.materialName)} /></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/** One sub-group as BMES > Make Model Group shows it: its name, how many line shifts, the shift groups, the sub-groups under it. */
function SubGroupCard(props: { node: SubGroupNode }) {
  const n = props.node;
  return (
    <div className="wk-bg-sub">
      <div className="wk-bg-row">
        <span className="wk-bg-badge">Sub Group {n.depth}</span><b>{n.name}</b>
        <span className="muted small wk-bg-count">{shiftCount(n)} shifts</span>
      </div>
      {n.lineShifts.map((ls) => <div key={ls} className="wk-bg-shift"><span className="wk-bg-sbadge">Shift Group</span>{ls}</div>)}
      {n.children.map((c) => <SubGroupCard key={c.name} node={c} />)}
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
      {tree.direct.map((ls) => <div key={ls} className="wk-bg-shift"><span className="wk-bg-sbadge">Shift Group</span>{ls}</div>)}
      {tree.nodes.map((n) => <SubGroupCard key={n.name} node={n} />)}
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
      <button className="btn sm primary" onClick={() => props.onLink(x.material)}>Link</button>
    </div>
  );
  return (
    <div className="wk-bg-picker">
      <div className="muted small">Pick the BMES &gt; Make Model Group model of {props.modelName}, then Save the model information.</div>
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
      {groups.filter((g) => list.some((x) => x.group === g)).map((g) => (
        <div key={g} className="wk-bg-group">
          <div className="wk-bg-gname">{g}</div>
          {list.filter((x) => x.group === g).map(row)}
        </div>
      ))}
      {!list.length && <div className="muted">No BMES model matches.</div>}
      <div className="wk-bg-foot">
        {props.onCancel && <button className="btn sm" onClick={props.onCancel}>Cancel</button>}
        <button className="wk-link" onClick={props.onManual}>Not in BMES Make Model Group? Enter the processes, NG items and materials by hand</button>
      </div>
    </div>
  );
}

/**
 * BMES model group of the model (BMES > Make Model Group), shown inside Model information as that page shows it: found
 * by the model name or picked from the Make Model Group list and saved with the model information, with its sub-groups,
 * shift groups and the BOM of each product. A model BMES does not have can get its group entered by hand.
 */
export default function BmesPanel(props: {
  modelName: string;
  /** Material saved in Model Info ("" = found by name). */
  saved: string;
  onPick: (material: string) => void;
  /** Model group entered by hand (Model Info), shown when it is picked or BMES has no group for the model. */
  manual: ManualGroupData;
  onManual: (g: ManualGroupData) => void;
  onError: (e: unknown) => void;
}) {
  const [materials, setMaterials] = useState<BmesMaterial[] | null>(null);
  const [product, setProduct] = useState("");
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
  const shown = m && m.products.includes(product) ? product : m?.products[0] ?? "";
  const hasManual = props.manual.processes.length > 0 || props.manual.materials.length > 0;
  const link = (material: string) => { props.onPick(material); setPicking(false); };

  let body;
  if (materials === null) body = <div className="muted">Loading the BMES model groups…</div>;
  else if (picking || (!m && !isManual && !hasManual)) {
    body = <GroupPicker modelName={props.modelName} materials={materials} current={m?.material ?? ""} onLink={link}
      onManual={() => link(MANUAL)} onCancel={picking ? () => setPicking(false) : undefined} />;
  } else if (!m) {
    body = <ManualGroup modelName={props.modelName} value={props.manual} onChange={props.onManual} onError={props.onError} />;
  } else {
    body = (
      <>
        <LinkedGroup m={m} />
        <div className="wk-bmes-bomh">
          <h3>BOM</h3>
          {m.products.map((p) => (
            <button key={p} className={"wk-chip" + (p === shown ? " on" : "")} onClick={() => setProduct(p)}>{p}</button>
          ))}
        </div>
        {shown ? <BomTable key={shown} product={shown} onError={props.onError} /> : <div className="muted">No product in this model.</div>}
      </>
    );
  }

  return (
    <div className="wk-bmes">
      <div className="wk-bmes-h">
        <h3>BMES model group</h3>
        {materials !== null && !picking && (m || isManual || hasManual) && (
          <span className="wk-bmes-pick">
            {m ? <>{m.group} › <b>{m.material}</b> · {found?.by === "saved" ? "linked in Model Info" : "same name"}</> : "Entered by hand"}
            <button className="btn sm" onClick={() => setPicking(true)}>Change link</button>
          </span>
        )}
      </div>
      {body}
    </div>
  );
}
