// Model group entered by hand in Model Info for a model that BMES > Make Model Group does not have: its processes,
// the NG items of each, the materials each uses per unit and the material names. Kept in the model information JSON.
import type { InstructionDoc } from "./workInstruction";

export const MANUAL = "__manual";
export const MANUAL_TYPES = ["SUB", "MAIN", "FUNCTION", "VISUAL"] as const;
export type ManualType = (typeof MANUAL_TYPES)[number];

export interface ManualUse {
  material: string;
  /** Pieces used per 1 unit going into the process. */
  qtyPerUnit: number;
}

export interface ManualProcess {
  id: string;
  type: ManualType;
  name: string;
  localName: string;
  ngItems: string[];
  materials: ManualUse[];
}

export interface ManualGroup {
  materials: string[];
  processes: ManualProcess[];
}

export const emptyManual = (): ManualGroup => ({ materials: [], processes: [] });

const key = (s: string) => s.trim().toLowerCase();
const newId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

export const blankProcess = (type: ManualType = "SUB"): ManualProcess =>
  ({ id: newId(), type, name: "", localName: "", ngItems: [], materials: [] });

/** NG items typed one per line: trimmed, empty lines and repeats (any case) dropped. */
export function ngFromText(text: string): string[] {
  const out: string[] = [];
  for (const line of text.split(/\r?\n/)) {
    const t = line.trim();
    if (t && !out.some((x) => key(x) === key(t))) out.push(t);
  }
  return out;
}

/** What is wrong with a process before it is kept, or "". */
export function processProblem(p: ManualProcess, materials: string[]): string {
  if (!p.name.trim()) return "Process name is required.";
  for (const u of p.materials) {
    if (!materials.some((m) => key(m) === key(u.material))) return `Material ${u.material || "(none)"} is not in the material list.`;
    if (!(u.qtyPerUnit > 0)) return `Qty per unit of ${u.material} must be more than 0.`;
  }
  const names = p.materials.map((u) => key(u.material));
  if (new Set(names).size !== names.length) return "A material is listed twice in this process.";
  return "";
}

/** The process cleaned for saving: trimmed names, NG items without repeats. */
export const cleanProcess = (p: ManualProcess): ManualProcess =>
  ({ ...p, name: p.name.trim(), localName: p.localName.trim(), ngItems: ngFromText(p.ngItems.join("\n")),
     materials: p.materials.map((u) => ({ material: u.material.trim(), qtyPerUnit: u.qtyPerUnit })) });

export function addMaterial(g: ManualGroup, name: string): ManualGroup {
  const n = name.trim();
  if (!n) throw new Error("Type a material name.");
  if (g.materials.some((m) => key(m) === key(n))) throw new Error(`${n} is already in the list.`);
  return { ...g, materials: [...g.materials, n] };
}

/**
 * The level-1 materials of a BOM (the product's direct materials) the list does not have yet, each once, in BOM order.
 * Input: the list and the BOM rows (level 1 = direct material). Output: the names to add.
 */
export function bomMaterials(g: ManualGroup, rows: { level: number; materialName: string; materialCode: string }[]): string[] {
  const out: string[] = [];
  for (const r of rows) {
    if (r.level !== 1) continue;
    const n = (r.materialName || r.materialCode).trim();
    if (n && !g.materials.some((m) => key(m) === key(n)) && !out.some((x) => key(x) === key(n))) out.push(n);
  }
  return out;
}

/** Adds these names to the end of the material list (ones already there left out). */
export const addMaterials = (g: ManualGroup, names: string[]): ManualGroup =>
  names.reduce((cur, n) => (n.trim() && !cur.materials.some((m) => key(m) === key(n)) ? { ...cur, materials: [...cur.materials, n.trim()] } : cur), g);

/** Removes a material from the list and from every process that used it. */
export const deleteMaterial = (g: ManualGroup, name: string): ManualGroup => ({
  materials: g.materials.filter((m) => key(m) !== key(name)),
  processes: g.processes.map((p) => ({ ...p, materials: p.materials.filter((u) => key(u.material) !== key(name)) })),
});

/** The names a work instruction gives the glue (bond) of a process in its tools or parts (Keo · SL233LF), as the server does. */
const GLUE_NAMES = ["keo", "keo dán", "bond", "glue", "adhesive", "epoxy", "본드", "접착제"];
/** A material named as glue by itself (Glue 233, Epoxy A). */
const GLUE_WORD = /\b(?:glue|bond|keo|adhesive|epoxy)\b|본드|접착제/i;
/** A BOM unit a glue is counted in (weight or volume), where a part is counted in pieces (PC). */
const GLUE_UNITS = ["G", "KG", "MG", "ML", "L"];

/** One material written another way (case, spaces, hyphens) is the same: SL-233LF is SL233LF. */
export const materialKey = (s: string) => s.replace(/[\s-]/g, "").toUpperCase();

/** The glue of the Work Instructions (material keys): the product in the note of a tool or part named as glue. */
export function glueOf(doc: Pick<InstructionDoc, "processes"> | null): Set<string> {
  const glue = new Set<string>();
  const add = (name: string, product: string) => {
    if (!GLUE_NAMES.includes(name.trim().toLowerCase())) return;
    const code = product.split("(")[0].trim(); // "SL233LF ( Sử dụng trong vòng 24 giờ)" → SL233LF
    if (code) glue.add(materialKey(code));
  };
  for (const p of doc?.processes ?? []) {
    p.tools.forEach((t) => add(t.name, t.note));
    p.parts.forEach((x) => add(x.name, x.spec));
  }
  return glue;
}

/** A bond: glue the Work Instructions name ("3777 (PUR)" goes with 3777), or named as glue by itself. */
export function isBond(name: string, glue: Set<string>): boolean {
  const k = materialKey(name);
  return glue.has(k) || [...glue].some((g) => k.startsWith(g + "(")) || GLUE_WORD.test(name);
}

/** The direct BOM materials counted by weight or volume (SL233LF in G): glue, not parts. */
export const bomGlue = (rows: { level: number; materialName: string; materialCode: string; usageUnit: string }[]): string[] =>
  rows.filter((r) => r.level === 1 && GLUE_UNITS.includes(r.usageUnit.trim().toUpperCase()))
    .map((r) => (r.materialName || r.materialCode).trim()).filter(Boolean);

/** Removes these materials from the list and from every process that used them. */
export const deleteMaterials = (g: ManualGroup, names: string[]): ManualGroup => names.reduce(deleteMaterial, g);

/** How many processes use each material. */
export const usage = (g: ManualGroup, name: string) => g.processes.filter((p) => p.materials.some((u) => key(u.material) === key(name))).length;

/**
 * Listed materials put into a process (dragged onto it) at 1 per unit, in list order; a material the process already
 * uses is skipped, and the group is unchanged if nothing was added.
 */
export function putMaterial(g: ManualGroup, processId: string, names: string[]): ManualGroup {
  const p = g.processes.find((x) => x.id === processId);
  if (!p) return g;
  const add = g.materials.filter((m) => names.some((n) => key(n) === key(m)) && !p.materials.some((u) => key(u.material) === key(m)));
  if (!add.length) return g;
  const next = { ...p, materials: [...p.materials, ...add.map((material) => ({ material, qtyPerUnit: 1 }))] };
  return { ...g, processes: g.processes.map((x) => (x.id === processId ? next : x)) };
}

export function saveProcess(g: ManualGroup, p: ManualProcess): ManualGroup {
  const clean = cleanProcess(p);
  const at = g.processes.findIndex((x) => x.id === p.id);
  return { ...g, processes: at < 0 ? [...g.processes, clean] : g.processes.map((x) => (x.id === p.id ? clean : x)) };
}

/** A process changed in place in the table, kept as typed (not trimmed) so typing is not disturbed; other processes stay. */
export const editProcess = (g: ManualGroup, p: ManualProcess): ManualGroup =>
  ({ ...g, processes: g.processes.map((x) => (x.id === p.id ? p : x)) });

export const deleteProcess =(g: ManualGroup, id: string): ManualGroup => ({ ...g, processes: g.processes.filter((p) => p.id !== id) });

export function moveProcess(g: ManualGroup, id: string, by: -1 | 1): ManualGroup {
  const i = g.processes.findIndex((p) => p.id === id);
  const j = i + by;
  if (i < 0 || j < 0 || j >= g.processes.length) return g;
  const list = [...g.processes];
  [list[i], list[j]] = [list[j], list[i]];
  return { ...g, processes: list };
}

/**
 * The group AI answered, read leniently, with the names of the processes it added (no id of the current group) and
 * removed. A material a process uses but the list does not have is added to the list.
 */
export function aiAnswer(g: ManualGroup, answer: unknown): { group: ManualGroup; added: string[]; removed: string[] } {
  const next = readManual(answer);
  const materials = [...next.materials];
  for (const u of next.processes.flatMap((p) => p.materials)) {
    if (u.material.trim() && !materials.some((m) => key(m) === key(u.material))) materials.push(u.material.trim());
  }
  const processes = next.processes.map(cleanProcess);
  const was = new Set(g.processes.map((p) => p.id));
  const now = new Set(processes.map((p) => p.id));
  return {
    group: { materials, processes },
    added: processes.filter((p) => !was.has(p.id)).map((p) => p.name),
    removed: g.processes.filter((p) => !now.has(p.id)).map((p) => p.name),
  };
}

export const isEmptyManual = (g: ManualGroup) => !g.processes.length && !g.materials.length;

/** The process tables of the sub-groups of a temporary model group (info.subManual), by sub-group id, read leniently. */
export function readSubManual(v: unknown): Record<string, ManualGroup> {
  const out: Record<string, ManualGroup> = {};
  if (v && typeof v === "object" && !Array.isArray(v)) for (const [id, g] of Object.entries(v)) out[id] = readManual(g);
  return out;
}

/**
 * Which process table is shown ("" = the model's, else a sub-group id). The model's table is offered while there is no
 * sub-group or it already has something; the one picked is kept while it is offered, else the first offered.
 */
export function shownTable(picked: string | null, subIds: string[], modelTable: ManualGroup): { at: string; model: boolean } {
  const model = !subIds.length || !isEmptyManual(modelTable);
  const offered = [...(model ? [""] : []), ...subIds];
  return { at: picked !== null && offered.includes(picked) ? picked : offered[0], model };
}

/** The tables of the sub-groups still there (a removed sub-group's table goes with it); the same object if none went. */
export function keepSubTables(t: Record<string, ManualGroup>, ids: string[]): Record<string, ManualGroup> {
  const gone = Object.keys(t).filter((id) => !ids.includes(id));
  if (!gone.length) return t;
  return Object.fromEntries(Object.entries(t).filter(([id]) => ids.includes(id)));
}

/** A saved group read leniently (older or hand-edited JSON). */
export function readManual(v: unknown): ManualGroup {
  const g = (v && typeof v === "object" ? v : {}) as Partial<ManualGroup>;
  const strings = (a: unknown) => (Array.isArray(a) ? a.filter((x): x is string => typeof x === "string") : []);
  return {
    materials: strings(g.materials),
    processes: (Array.isArray(g.processes) ? g.processes : []).map((p) => ({
      id: String(p?.id || newId()),
      type: (MANUAL_TYPES as readonly string[]).includes(p?.type) ? p.type : "SUB",
      name: String(p?.name ?? ""),
      localName: String(p?.localName ?? ""),
      ngItems: strings(p?.ngItems),
      materials: (Array.isArray(p?.materials) ? p.materials : [])
        .map((u) => ({ material: String(u?.material ?? ""), qtyPerUnit: Number(u?.qtyPerUnit) || 0 })),
    })),
  };
}
