// Temporary model group of a model BMES does not have yet (a development model): the Model › Sub Group › Shift Group
// tree of BMES > Make Model Group, made in Model Info by hand and kept in the model information JSON (info.tempGroup)
// until the model is in BMES and its real group can be linked.

export interface TempSub {
  id: string;
  name: string;
  /** Shift groups (line shifts, MATERIALNAME_LINE) put in this sub-group. */
  lineShifts: string[];
  children: TempSub[];
}

export interface TempGroup {
  /** Product group of the model (MSM, TIM, …), as BMES > Make Model Group groups the models. */
  group: string;
  /** Shift groups put on the model itself, outside any sub-group. */
  lineShifts: string[];
  subGroups: TempSub[];
}

export const emptyTempGroup = (): TempGroup => ({ group: "", lineShifts: [], subGroups: [] });

const key = (s: string) => s.trim().toLowerCase();
const newId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);

/** Every shift group of the group, the model's first, then the sub-groups' top to bottom. */
export function allShifts(g: TempGroup): string[] {
  const walk = (list: TempSub[]): string[] => list.flatMap((s) => [...s.lineShifts, ...walk(s.children)]);
  return [...g.lineShifts, ...walk(g.subGroups)];
}

/** Sub-groups and shift groups in the group (a sub-group counted at any depth). */
export function tempCounts(g: TempGroup): { subs: number; shifts: number } {
  const subs = (list: TempSub[]): number => list.reduce((n, s) => n + 1 + subs(s.children), 0);
  return { subs: subs(g.subGroups), shifts: allShifts(g).length };
}

export const hasTempGroup = (g: TempGroup) => !!g.group.trim() || tempCounts(g).subs > 0 || g.lineShifts.length > 0;

/** Shift groups of a sub-group and every sub-group under it. */
export const subShiftCount = (s: TempSub): number => s.lineShifts.length + s.children.reduce((n, c) => n + subShiftCount(c), 0);

/** The sub-groups with the one of this id changed (at any depth). */
function mapSub(list: TempSub[], id: string, f: (s: TempSub) => TempSub | null): TempSub[] {
  return list.flatMap((s) => {
    if (s.id === id) { const n = f(s); return n ? [n] : []; }
    return [{ ...s, children: mapSub(s.children, id, f) }];
  });
}

function findSub(list: TempSub[], id: string): TempSub | null {
  for (const s of list) {
    if (s.id === id) return s;
    const c = findSub(s.children, id);
    if (c) return c;
  }
  return null;
}

/** The sub-groups next to a new or renamed one: the top ones, or the children of its parent. */
const siblings = (g: TempGroup, parentId: string | null) => (parentId ? findSub(g.subGroups, parentId)?.children ?? [] : g.subGroups);

/** A sub-group added at the end of the model (parentId null) or of a sub-group; the name must be new among its siblings. */
export function addSub(g: TempGroup, parentId: string | null, name: string): TempGroup {
  const n = name.trim();
  if (!n) throw new Error("Type a sub group name.");
  if (siblings(g, parentId).some((s) => key(s.name) === key(n))) throw new Error(`${n} is already a sub group here.`);
  const sub: TempSub = { id: newId(), name: n, lineShifts: [], children: [] };
  if (!parentId) return { ...g, subGroups: [...g.subGroups, sub] };
  return { ...g, subGroups: mapSub(g.subGroups, parentId, (p) => ({ ...p, children: [...p.children, sub] })) };
}

/** The sub-group renamed; an empty name or one a sibling already has is refused. */
export function renameSub(g: TempGroup, id: string, name: string): TempGroup {
  const n = name.trim();
  if (!n) throw new Error("Type a sub group name.");
  const parent = (list: TempSub[], up: string | null): string | null | undefined => {
    for (const s of list) {
      if (s.id === id) return up;
      const p = parent(s.children, s.id);
      if (p !== undefined) return p;
    }
    return undefined;
  };
  const up = parent(g.subGroups, null);
  if (up === undefined) return g;
  if (siblings(g, up).some((s) => s.id !== id && key(s.name) === key(n))) throw new Error(`${n} is already a sub group here.`);
  return { ...g, subGroups: mapSub(g.subGroups, id, (s) => ({ ...s, name: n })) };
}

/** The sub-group removed with its shift groups and the sub-groups under it. */
export const removeSub = (g: TempGroup, id: string): TempGroup => ({ ...g, subGroups: mapSub(g.subGroups, id, () => null) });

/**
 * A shift group put on the model (at null) or a sub-group. A shift group sits in one place only, so one already in the
 * group (any case) is refused.
 */
export function addShift(g: TempGroup, at: string | null, name: string): TempGroup {
  const n = name.trim();
  if (!n) throw new Error("Type a shift group name.");
  if (allShifts(g).some((x) => key(x) === key(n))) throw new Error(`${n} is already in this model group.`);
  if (!at) return { ...g, lineShifts: [...g.lineShifts, n] };
  if (!findSub(g.subGroups, at)) return g;
  return { ...g, subGroups: mapSub(g.subGroups, at, (s) => ({ ...s, lineShifts: [...s.lineShifts, n] })) };
}

/** The shift group taken out of wherever it is. */
export function removeShift(g: TempGroup, name: string): TempGroup {
  const drop = (list: string[]) => list.filter((x) => key(x) !== key(name));
  const walk = (list: TempSub[]): TempSub[] => list.map((s) => ({ ...s, lineShifts: drop(s.lineShifts), children: walk(s.children) }));
  return { ...g, lineShifts: drop(g.lineShifts), subGroups: walk(g.subGroups) };
}

/** A saved group read leniently (older or hand-edited JSON). */
export function readTempGroup(v: unknown): TempGroup {
  const g = (v && typeof v === "object" ? v : {}) as Record<string, unknown>;
  const strings = (a: unknown) => (Array.isArray(a) ? a.filter((x): x is string => typeof x === "string" && !!x.trim()).map((x) => x.trim()) : []);
  const subs = (a: unknown): TempSub[] => (Array.isArray(a) ? a : [])
    .filter((s): s is Record<string, unknown> => !!s && typeof s === "object")
    .map((s) => ({ id: typeof s.id === "string" && s.id ? s.id : newId(), name: String(s.name ?? ""), lineShifts: strings(s.lineShifts), children: subs(s.children) }));
  return { group: typeof g.group === "string" ? g.group : "", lineShifts: strings(g.lineShifts), subGroups: subs(g.subGroups) };
}
