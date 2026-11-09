import { useEffect, useState } from "react";
import { bmesApi, matchMaterial, qty, visibleBom } from "./bmes";
import type { BmesMaterial, BomResult } from "./bmes";

/** BOM of one product: from the local cache when there is one, or loaded from BMES on request. */
function BomTable(props: { product: string; onError: (e: unknown) => void }) {
  const { product } = props;
  const [bom, setBom] = useState<BomResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [closed, setClosed] = useState<Set<number>>(new Set());

  useEffect(() => {
    let live = true;
    setBom(null);
    setClosed(new Set());
    bmesApi.bom(product).then((b) => { if (live) setBom(b); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [product]);

  const load = async () => {
    setBusy(true);
    try { setBom(await bmesApi.fetchBom(product)); setClosed(new Set()); } catch (e) { props.onError(e); } finally { setBusy(false); }
  };
  const toggle = (i: number) => setClosed((s) => { const n = new Set(s); if (n.has(i)) n.delete(i); else n.add(i); return n; });

  return (
    <div className="wk-bom">
      <div className="wk-bom-h">
        <b>{product}</b>
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
      {!!bom?.rows.length && (
        <div className="pi-wrap">
          <table className="wk-it wk-bom-t">
            <thead>
              <tr><th>Level</th><th className="l">Material</th><th>Code</th><th>Qty</th><th>Unit</th></tr>
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
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

/**
 * BMES model group of the model (BMES > Make Model Group): found by the model name or picked here and saved with the
 * model information, with its sub-groups, line shifts and the BOM of each product.
 */
export default function BmesPanel(props: {
  modelName: string;
  /** Material saved in Model Info ("" = found by name). */
  saved: string;
  onPick: (material: string) => void;
  onError: (e: unknown) => void;
}) {
  const [materials, setMaterials] = useState<BmesMaterial[] | null>(null);
  const [product, setProduct] = useState("");

  useEffect(() => {
    let live = true;
    bmesApi.materials().then((m) => { if (live) setMaterials(m); }).catch(props.onError);
    return () => { live = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const found = materials ? matchMaterial(props.modelName, materials, props.saved) : null;
  const m = found?.material ?? null;
  const shown = m && m.products.includes(product) ? product : m?.products[0] ?? "";
  const byName = materials ? matchMaterial(props.modelName, materials).material : null;
  const groups = [...new Set((materials ?? []).map((x) => x.group))];

  return (
    <section className="card pi-sec">
      <div className="pi-sec-h">
        <h2>BMES model group</h2>
        <div className="pi-legend">
          <label className="wk-bmes-pick">
            Linked material{" "}
            <select className="inp sm" value={props.saved} disabled={!materials} onChange={(e) => props.onPick(e.target.value)}>
              <option value="">{byName ? `Same name: ${byName.material}` : "Same name (none found)"}</option>
              {groups.map((g) => (
                <optgroup key={g} label={g}>
                  {materials!.filter((x) => x.group === g).map((x) => <option key={x.material} value={x.material}>{x.material}</option>)}
                </optgroup>
              ))}
            </select>
          </label>
        </div>
      </div>
      <div className="pi-sec-b">
        {materials === null ? <div className="muted">Loading the BMES model groups…</div>
          : !m ? (
            <div className="muted">
              No BMES model group has a material named like {props.modelName}. Pick the material above, or add it in BMES &gt; Make Model Group.
            </div>
          ) : (
            <>
              <div className="wk-bmes-path">
                <span className="wk-chip on">{m.group}</span><span className="muted">›</span><b>{m.material}</b>
                {m.productGroup && m.productGroup !== m.group && <span className="muted small">Product group {m.productGroup}</span>}
                <span className="muted small">{found?.by === "saved" ? "picked in Model Info" : "found by the model name"}</span>
              </div>
              <table className="wk-it wk-bmes-sub">
                <thead><tr><th className="l">Sub-group</th><th className="l">Line shifts</th></tr></thead>
                <tbody>
                  {m.subGroups.map((s) => (
                    <tr key={s.path}><td className="l">{s.path}</td><td className="l wk-bmes-ls">{s.lineShifts.join(", ")}</td></tr>
                  ))}
                  {!m.subGroups.length && <tr><td colSpan={2} className="wk-it-empty">No line shift in this material.</td></tr>}
                </tbody>
              </table>
              <div className="wk-bmes-bomh">
                <h3>BOM</h3>
                {m.products.map((p) => (
                  <button key={p} className={"wk-chip" + (p === shown ? " on" : "")} onClick={() => setProduct(p)}>{p}</button>
                ))}
              </div>
              {shown ? <BomTable key={shown} product={shown} onError={props.onError} /> : <div className="muted">No product in this material.</div>}
            </>
          )}
      </div>
    </section>
  );
}
