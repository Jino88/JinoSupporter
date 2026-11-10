// Model group entered by hand in Model Info for a model that BMES > Make Model Group does not have: its processes,
// the NG items of each, the materials each uses per unit and the material names. Kept in the model information JSON.
import type { ModelSetup } from "../productionInput/types";

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

/** Removes a material from the list and from every process that used it. */
export const deleteMaterial = (g: ManualGroup, name: string): ManualGroup => ({
  materials: g.materials.filter((m) => key(m) !== key(name)),
  processes: g.processes.map((p) => ({ ...p, materials: p.materials.filter((u) => key(u.material) !== key(name)) })),
});

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

export const deleteProcess = (g: ManualGroup, id: string): ManualGroup => ({ ...g, processes: g.processes.filter((p) => p.id !== id) });

export function moveProcess(g: ManualGroup, id: string, by: -1 | 1): ManualGroup {
  const i = g.processes.findIndex((p) => p.id === id);
  const j = i + by;
  if (i < 0 || j < 0 || j >= g.processes.length) return g;
  const list = [...g.processes];
  [list[i], list[j]] = [list[j], list[i]];
  return { ...g, processes: list };
}

/** The processes, NG items and materials of a Production Input model setup. */
export const fromProductionInput = (s: ModelSetup): ManualGroup => ({
  materials: [...s.materials],
  processes: s.processes.map((p) => ({
    id: newId(), type: p.type, name: p.name, localName: p.localName ?? "",
    ngItems: p.ngItems.map((n) => n.name), materials: p.materials.map((u) => ({ material: u.material, qtyPerUnit: u.qtyPerUnit })),
  })),
});

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
