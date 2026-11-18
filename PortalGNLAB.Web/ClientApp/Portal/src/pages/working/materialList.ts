// Material tab of a model: the materials of each group of the model (a product of its BMES model group, a sub-group of
// its temporary group, or the model itself), typed by hand or loaded from the group's BOM, each with the mold it is made
// in (mass production or sample). Kept in the model information JSON (info.materials).
import { qty as bomQty, type BmesMaterial, type BomRow } from "./bmes";
import { subPaths, type TempGroup } from "./tempGroup";

/** The mold a material is made in: "" = not set. */
export type MoldKind = "" | "mass" | "sample";
export const MOLD_KINDS: { key: MoldKind; label: string }[] = [
  { key: "", label: "—" },
  { key: "mass", label: "Mass production mold" },
  { key: "sample", label: "Sample mold" },
];

export interface MaterialItem {
  id: string;
  code: string;
  name: string;
  /** Usage per unit, as typed or as the BOM gives it. */
  qty: string;
  unit: string;
  mold: MoldKind;
  note: string;
}

/** A group of the model and its materials; `label` is kept so a group no longer in BMES still shows its name. */
export interface MaterialGroup {
  label: string;
  rows: MaterialItem[];
}

/** Group key → its materials. */
export type MaterialList = Record<string, MaterialGroup>;

/** A group the model has: where its BOM is read from (the product name in BMES). */
export interface GroupDef {
  key: string;
  label: string;
  product: string;
}

const newId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const norm = (s: string) => s.trim().toLowerCase();

export const blankMaterial = (): MaterialItem => ({ id: newId(), code: "", name: "", qty: "", unit: "", mold: "", note: "" });

/**
 * The groups of a model: each product of its BMES model group, else each sub-group of its temporary group (its BOM by the
 * last name of the path), else the model itself; then the saved groups not in that list any more, so nothing typed is hidden.
 */
export function materialGroups(model: string, bmes: BmesMaterial | null, temp: TempGroup, saved: MaterialList): GroupDef[] {
  let out: GroupDef[];
  if (bmes?.products.length) out = bmes.products.map((p) => ({ key: `p:${p}`, label: p, product: p }));
  else {
    const subs = subPaths(temp);
    out = subs.length ? subs.map((s) => ({ key: `s:${s.id}`, label: s.path, product: s.path.split(" › ").pop() ?? s.path }))
      : [{ key: "m", label: model, product: model }];
  }
  for (const [key, g] of Object.entries(saved)) {
    if (!out.some((d) => d.key === key) && g.rows.length) out.push({ key, label: g.label || key, product: g.label || model });
  }
  return out;
}

/** The materials of a BOM the group does not have yet (same code, or same name when there is no code), in BOM order. */
export function bomMaterialItems(rows: MaterialItem[], bom: BomRow[], allLevels = false): MaterialItem[] {
  const out: MaterialItem[] = [];
  const has = (list: MaterialItem[], code: string, name: string) =>
    list.some((m) => (code && norm(m.code) === norm(code)) || (!code && name && norm(m.name) === norm(name)));
  for (const r of bom) {
    if (!allLevels && r.level !== 1) continue;
    const code = r.materialCode.trim(), name = r.materialName.trim();
    if (!code && !name) continue;
    if (has(rows, code, name) || has(out, code, name)) continue;
    out.push({ ...blankMaterial(), code, name, qty: bomQty(r.usageQty), unit: r.usageUnit?.trim() ?? "" });
  }
  return out;
}

/** What is wrong with the rows before saving, or "": a row needs a code or a name, and a code once per group. */
export function materialProblem(rows: MaterialItem[]): string {
  const seen = new Set<string>();
  for (const [i, r] of rows.entries()) {
    if (!r.code.trim() && !r.name.trim()) return `Row ${i + 1} needs a code or a name.`;
    const c = norm(r.code);
    if (c && seen.has(c)) return `The code ${r.code.trim()} is listed twice.`;
    if (c) seen.add(c);
  }
  return "";
}

/** The list made ready to save: trimmed texts, empty groups left out. */
export function cleanMaterials(list: MaterialList): MaterialList {
  const out: MaterialList = {};
  for (const [key, g] of Object.entries(list)) {
    const rows = g.rows.map((r) => ({ ...r, code: r.code.trim(), name: r.name.trim(), qty: r.qty.trim(), unit: r.unit.trim(), note: r.note.trim() }));
    if (rows.length) out[key] = { label: g.label, rows };
  }
  return out;
}

/** How many rows use each mold kind. */
export const moldCounts = (rows: MaterialItem[]) =>
  ({ mass: rows.filter((r) => r.mold === "mass").length, sample: rows.filter((r) => r.mold === "sample").length });

const text = (v: unknown) => (typeof v === "string" ? v : typeof v === "number" ? String(v) : "");

/** info.materials as saved (anything missing or of another shape becomes empty). */
export function readMaterials(v: unknown): MaterialList {
  const out: MaterialList = {};
  if (!v || typeof v !== "object" || Array.isArray(v)) return out;
  for (const [key, g] of Object.entries(v as Record<string, unknown>)) {
    if (!g || typeof g !== "object") continue;
    const raw = (g as { label?: unknown; rows?: unknown });
    const rows = (Array.isArray(raw.rows) ? raw.rows : []).filter((r): r is Record<string, unknown> => !!r && typeof r === "object").map((r) => ({
      id: text(r.id) || newId(), code: text(r.code), name: text(r.name), qty: text(r.qty), unit: text(r.unit),
      mold: (r.mold === "mass" || r.mold === "sample" ? r.mold : "") as MoldKind, note: text(r.note),
    }));
    out[key] = { label: text(raw.label), rows };
  }
  return out;
}
