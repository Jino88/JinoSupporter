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

/**
 * BMES materials that look like the model when the name does not match: they share a number of 3 or more digits with it
 * (TIM-338 → ASSY 338 RA1, ASSY REAR-TAPE-338).
 */
export function suggestMaterials(model: string, materials: BmesMaterial[]): BmesMaterial[] {
  const nums = new Set(model.match(/\d{3,}/g) ?? []);
  if (!nums.size) return [];
  return materials.filter((m) => (m.material.match(/\d{3,}/g) ?? []).some((n) => nums.has(n)));
}

/** A sub-group of a BMES material as BMES > Make Model Group shows it: its name, depth (1 = Sub Group 1), line shifts. */
export interface SubGroupNode {
  name: string;
  depth: number;
  lineShifts: string[];
  children: SubGroupNode[];
}

/** The sub-groups ("A › B" paths) as the tree of Make Model Group; a path left empty puts its line shifts on the model. */
export function subGroupTree(subGroups: BmesMaterial["subGroups"]): { direct: string[]; nodes: SubGroupNode[] } {
  const direct: string[] = [];
  const nodes: SubGroupNode[] = [];
  for (const s of subGroups) {
    const parts = s.path.split(" › ").map((p) => p.trim()).filter(Boolean);
    if (!parts.length) { direct.push(...s.lineShifts); continue; }
    let level = nodes;
    let node: SubGroupNode | undefined;
    parts.forEach((name, i) => {
      node = level.find((n) => n.name === name);
      if (!node) { node = { name, depth: i + 1, lineShifts: [], children: [] }; level.push(node); }
      level = node.children;
    });
    node!.lineShifts.push(...s.lineShifts);
  }
  return { direct, nodes };
}

/** Line shifts of a sub-group and every sub-group under it. */
export const shiftCount = (n: SubGroupNode): number => n.lineShifts.length + n.children.reduce((s, c) => s + shiftCount(c), 0);

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

/** A drawing (PDM master row) of a material: PDRTX is the material name it belongs to. */
export interface PdmRow {
  pdmno: string;
  pdmvr: string;
  pdtnoTx: string;
  pdrtx: string;
  /** Create Date (ERDAT) and Mod. Date (ARDAT) of the drawing in BMES; absent from an older server. */
  erdat?: string;
  ardat?: string;
}

export const drawingApi = {
  /** Drawings of these material names, from the local PDM master (BOM & Drawing > Fetch PDM). */
  find: (names: string[]) => toolSend<{ rows: PdmRow[] }>("working/bmes/pdm", { names }).then((r) => r.rows),
  /** The drawing file, opened in a new tab (read from BMES PDM). */
  url: (d: PdmRow) => toolPath(`working/bmes/drawing?pdmno=${encodeURIComponent(d.pdmno)}&pdmvr=${encodeURIComponent(d.pdmvr || "0")}`),
};

/** Material name (case and spaces ignored) → its drawings, each PDM number and version once. */
export function drawingsByName(rows: PdmRow[]): Map<string, PdmRow[]> {
  const out = new Map<string, PdmRow[]>();
  for (const r of rows) {
    const key = norm(r.pdrtx);
    if (!key) continue;
    const list = out.get(key) ?? [];
    if (!list.some((x) => x.pdmno === r.pdmno && x.pdmvr === r.pdmvr)) list.push(r);
    out.set(key, list);
  }
  return out;
}

export const drawingsOf = (by: Map<string, PdmRow[]>, materialName: string) => by.get(norm(materialName)) ?? [];
