// BMES model groups and BOMs of a Working model (Endpoints/WorkingEndpoints.cs, /bmes). The groups are the ones made in
// BMES > Make Model Group; a BOM comes from the local BOM cache or, on request, from BMES (read only).
import { getJson, toolPath, toolSend } from "../../toolsApi";

/** A material of a BMES model group with the products (BOM names) its line shifts stand for. */
export interface BmesMaterial {
  group: string;
  productGroup: string;
  material: string;
  products: string[];
  subGroups: { path: string; lineShifts: string[] }[];
}

export interface BomRow {
  level: number;
  materialCode: string;
  materialName: string;
  usageQty: string | number | null;
  usageUnit: string;
  parentMaterialName: string;
  hasChildren: boolean;
}

export interface BomResult {
  product: string;
  /** "cache", "server", or "" when there is no BOM yet. */
  source: string;
  fetchedAt: string;
  workDate: string;
  rows: BomRow[];
  error: string;
}

export const bmesApi = {
  materials: () => getJson<{ materials: BmesMaterial[] }>(toolPath("working/bmes/groups")).then((r) => r.materials),
  bom: (product: string) => getJson<BomResult>(toolPath(`working/bmes/bom?product=${encodeURIComponent(product)}`)),
  fetchBom: (product: string) => toolSend<BomResult>("working/bmes/bom", { product }),
};

const norm = (s: string) => s.toUpperCase().replace(/\s+/g, "");
/** The name without the -ZZ / ZZ ending the model list adds (TIU-C11-18-ZZ and TIU-C11-18 are the same model). */
const core = (s: string) => norm(s).replace(/-?ZZ$/, "");

/**
 * The BMES material of a model: the one saved in Model Info, else the one with the same name (ignoring case, spaces
 * and a -ZZ ending).
 */
export function matchMaterial(model: string, materials: BmesMaterial[], saved = ""): { material: BmesMaterial | null; by: "saved" | "name" | "" } {
  if (saved) {
    const m = materials.find((x) => x.material === saved);
    if (m) return { material: m, by: "saved" };
  }
  const exact = materials.find((x) => norm(x.material) === norm(model));
  if (exact) return { material: exact, by: "name" };
  const near = materials.find((x) => core(x.material) === core(model));
  return near ? { material: near, by: "name" } : { material: null, by: "" };
}

/** 1000.000000 → 1000, 0.250000 → 0.25 (BMES keeps six decimals); no quantity → "". */
export function qty(v: string | number | null): string {
  if (v === null) return "";
  const n = typeof v === "number" ? v : Number(String(v).trim());
  if (!Number.isFinite(n) || String(v).trim() === "") return String(v);
  return String(Number(n.toFixed(6)));
}

/** Rows whose parents are all open: a row is hidden while one of the rows above it in its branch is closed. */
export function visibleBom(rows: BomRow[], closed: Set<number>): { row: BomRow; index: number }[] {
  const out: { row: BomRow; index: number }[] = [];
  let hideBelow = Infinity; // rows deeper than this level are inside a closed branch
  rows.forEach((row, index) => {
    if (row.level > hideBelow) return;
    hideBelow = closed.has(index) ? row.level : Infinity;
    out.push({ row, index });
  });
  return out;
}
