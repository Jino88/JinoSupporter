import type { MaterialRow, ModelGroup, SubGroup } from "../../bmesApi";

export const PRODUCT_GROUPS = ["SPK", "UNIT", "MODULE", "TWS", "ETC"];

export interface SubVm { id: string; name: string; lineShifts: string[]; subGroups: SubVm[] }
export interface MidVm { id: string; material: string; subGroups: SubVm[] }
export interface GroupVm { id: string; name: string; productGroup: string; midGroups: MidVm[] }

/** The editable groups plus what the next click adds to (group → model → sub group). */
export interface Editor {
  groups: GroupVm[];
  activeGroup: number;
  activeMid: number;
  activeSub: string | null;
}

let counter = 0;
const uid = () => `n${++counter}`;
const newSub = (name = ""): SubVm => ({ id: uid(), name, lineShifts: [], subGroups: [] });
const newMid = (material = ""): MidVm => ({ id: uid(), material, subGroups: [newSub()] });
const newGroup = (name: string): GroupVm => ({ id: uid(), name, productGroup: "ETC", midGroups: [] });

export const initialEditor = (): Editor => ({ groups: [newGroup("Group 1")], activeGroup: 0, activeMid: -1, activeSub: null });

const subFromApi = (s: SubGroup): SubVm => ({ id: uid(), name: s.name ?? "", lineShifts: [...s.lineShifts], subGroups: s.subGroups.map(subFromApi) });
const subToApi = (s: SubVm): SubGroup => ({ name: s.name, lineShifts: [...s.lineShifts], subGroups: s.subGroups.map(subToApi) });

/** Editor for the saved groups; the default single empty group when nothing is saved. */
export function fromApi(groups: ModelGroup[]): Editor {
  if (groups.length === 0) return initialEditor();
  return {
    groups: groups.map((g) => ({
      id: uid(),
      name: g.name,
      productGroup: PRODUCT_GROUPS.includes(g.productGroup) ? g.productGroup : "ETC",
      midGroups: g.midGroups.map((m) => ({
        id: uid(),
        material: m.material,
        subGroups: m.subGroups.length > 0 ? m.subGroups.map(subFromApi) : [newSub()],
      })),
    })),
    activeGroup: 0,
    activeMid: -1,
    activeSub: null,
  };
}

/** Payload of Save: same shape the Blazor page builds (sort order = position). */
export function toApi(groups: GroupVm[]): ModelGroup[] {
  return groups.map((g) => ({
    name: g.name,
    productGroup: PRODUCT_GROUPS.includes(g.productGroup) ? g.productGroup : "ETC",
    midGroups: g.midGroups.map((m) => ({ material: m.material, subGroups: m.subGroups.map(subToApi) })),
  }));
}

export const totalLineShifts = (s: SubVm): number => s.lineShifts.length + s.subGroups.reduce((n, c) => n + totalLineShifts(c), 0);
export const midTotal = (m: MidVm): number => m.subGroups.reduce((n, s) => n + totalLineShifts(s), 0);

export function countSubs(subs: SubGroup[]): number {
  return subs.reduce((n, s) => n + 1 + countSubs(s.subGroups), 0);
}

const clone = (e: Editor): Editor => structuredClone(e);

function eachSub(subs: SubVm[], visit: (s: SubVm, parent: SubVm[]) => void) {
  for (const s of subs) {
    visit(s, subs);
    eachSub(s.subGroups, visit);
  }
}

function findSub(e: Editor, id: string): { sub: SubVm; list: SubVm[]; gi: number; mi: number } | null {
  for (let gi = 0; gi < e.groups.length; gi++)
    for (let mi = 0; mi < e.groups[gi].midGroups.length; mi++) {
      let found: { sub: SubVm; list: SubVm[] } | null = null;
      eachSub(e.groups[gi].midGroups[mi].subGroups, (s, list) => { if (s.id === id) found = { sub: s, list }; });
      if (found) return { ...(found as { sub: SubVm; list: SubVm[] }), gi, mi };
    }
  return null;
}

const contains = (parent: SubVm, id: string): boolean => parent.id === id || parent.subGroups.some((c) => contains(c, id));

export function activeMid(e: Editor): MidVm | null {
  return e.groups[e.activeGroup]?.midGroups[e.activeMid] ?? null;
}

export function addGroup(e: Editor): Editor {
  const next = clone(e);
  next.groups.push(newGroup(`Group ${next.groups.length + 1}`));
  next.activeGroup = next.groups.length - 1;
  return next;
}

export function setActiveGroup(e: Editor, idx: number): Editor {
  if (idx < 0 || idx >= e.groups.length) return e;
  const next = clone(e);
  if (next.activeGroup !== idx) { next.activeMid = -1; next.activeSub = null; }
  next.activeGroup = idx;
  return next;
}

export function removeGroup(e: Editor, idx: number): Editor {
  if (idx < 0 || idx >= e.groups.length) return e;
  const next = clone(e);
  next.groups.splice(idx, 1);
  if (next.groups.length === 0) { next.groups.push(newGroup("Group 1")); next.activeGroup = 0; }
  else if (next.activeGroup >= next.groups.length) next.activeGroup = next.groups.length - 1;
  return next;
}

export const clearGroups = (): Editor => initialEditor();

/** Nothing to clear: one group with no models. */
export const isCleared = (groups: GroupVm[]) => groups.length <= 1 && groups.every((g) => g.midGroups.length === 0);

export function renameGroup(e: Editor, idx: number, name: string): Editor {
  if (!e.groups[idx]) return e;
  const next = clone(e);
  next.groups[idx].name = name;
  return next;
}

export function setProductGroup(e: Editor, idx: number, pg: string): Editor {
  if (!e.groups[idx]) return e;
  const next = clone(e);
  next.groups[idx].productGroup = PRODUCT_GROUPS.includes(pg) ? pg : "ETC";
  return next;
}

export function addMid(e: Editor, gi: number): Editor {
  if (!e.groups[gi]) return e;
  const next = clone(e);
  const mid = newMid();
  next.groups[gi].midGroups.push(mid);
  next.activeGroup = gi;
  next.activeMid = next.groups[gi].midGroups.length - 1;
  next.activeSub = mid.subGroups[0].id;
  return next;
}

export function renameMid(e: Editor, gi: number, mi: number, material: string): Editor {
  if (!e.groups[gi]?.midGroups[mi]) return e;
  const next = clone(e);
  next.groups[gi].midGroups[mi].material = material;
  return next;
}

export function removeMid(e: Editor, gi: number, mi: number): Editor {
  if (!e.groups[gi]?.midGroups[mi]) return e;
  const next = clone(e);
  const grp = next.groups[gi];
  grp.midGroups.splice(mi, 1);
  if (next.activeGroup === gi && next.activeMid >= grp.midGroups.length) next.activeMid = grp.midGroups.length - 1;
  return next;
}

export function setActiveMid(e: Editor, gi: number, mi: number): Editor {
  if (!e.groups[gi]) return e;
  const out = clone(setActiveGroup(e, gi));
  if (out.activeMid !== mi) out.activeSub = null;
  out.activeMid = mi;
  return out;
}

/** Makes `id` the active sub group and moves the active group / model to where it lives. */
export function setActiveSub(e: Editor, id: string): Editor {
  const next = clone(e);
  next.activeSub = id;
  const at = findSub(next, id);
  if (at) { next.activeGroup = at.gi; next.activeMid = at.mi; }
  return next;
}

/** "+ Sub Group" inside a sub group (parentId) or at the top of the active model (null). */
export function addSubUnder(e: Editor, parentId: string | null): Editor {
  const next = clone(e);
  const list = parentId === null ? activeMid(next)?.subGroups : findSub(next, parentId)?.sub.subGroups;
  if (!list) return e;
  const sub = newSub(`Line ${list.length + 1}`);
  list.push(sub);
  return setActiveSub(next, sub.id);
}

export function addSubAtMidTop(e: Editor, gi: number, mi: number): Editor {
  if (!e.groups[gi]?.midGroups[mi]) return e;
  const next = clone(e);
  const mid = next.groups[gi].midGroups[mi];
  const sub = newSub(`Line ${mid.subGroups.length + 1}`);
  mid.subGroups.push(sub);
  next.activeGroup = gi;
  next.activeMid = mi;
  next.activeSub = sub.id;
  return next;
}

export function renameSub(e: Editor, id: string, name: string): Editor {
  const next = clone(e);
  const at = findSub(next, id);
  if (!at) return e;
  at.sub.name = name;
  return next;
}

/** Removes a sub group wherever it is. A model left without sub groups gets an empty default one. */
export function removeSub(e: Editor, id: string): Editor {
  const next = clone(e);
  const at = findSub(next, id);
  if (!at) return e;
  at.list.splice(at.list.findIndex((s) => s.id === id), 1);
  const mid = next.groups[at.gi].midGroups[at.mi];
  if (mid.subGroups.length === 0) mid.subGroups.push(newSub());
  if (next.activeSub === id) next.activeSub = null;
  return next;
}

export function removeLineShift(e: Editor, subId: string, lineShift: string): Editor {
  const next = clone(e);
  const at = findSub(next, subId);
  if (!at) return e;
  const i = at.sub.lineShifts.indexOf(lineShift);
  if (i >= 0) at.sub.lineShifts.splice(i, 1);
  return next;
}

/** A sub group can go into another one of the same model, never into itself or its own descendant. */
export function canMoveSub(e: Editor, sourceId: string, targetId: string): boolean {
  if (sourceId === targetId) return false;
  const source = findSub(e, sourceId);
  const target = findSub(e, targetId);
  if (!source || !target) return false;
  if (contains(source.sub, targetId)) return false;
  return source.gi === target.gi && source.mi === target.mi;
}

export const MOVE_ERROR = "Sub Group can only be dropped into another Sub Group under the same Model.";

const subName = (s: SubVm) => (s.name.trim() ? s.name : "Sub Group");

export function dropSub(e: Editor, sourceId: string, targetId: string): { editor: Editor; message: string; error: boolean } {
  if (!canMoveSub(e, sourceId, targetId)) return { editor: e, message: MOVE_ERROR, error: true };
  const next = clone(e);
  const source = findSub(next, sourceId)!;
  const target = findSub(next, targetId)!;
  source.list.splice(source.list.findIndex((s) => s.id === sourceId), 1);
  target.sub.subGroups.push(source.sub);
  return {
    editor: setActiveSub(next, sourceId),
    message: `Moved Sub Group '${subName(source.sub)}' into '${subName(target.sub)}'.`,
    error: false,
  };
}

/** Click on a Shift Group row: puts it into the active sub group, creating group / model / sub group as needed. */
export function addLineShift(e: Editor, lineShift: string, selectedMaktx: string): Editor {
  if (!lineShift) return e;
  const next = clone(e);
  if (next.groups.length === 0) { next.groups.push(newGroup("Group 1")); next.activeGroup = 0; }
  if (next.activeGroup < 0 || next.activeGroup >= next.groups.length) next.activeGroup = next.groups.length - 1;

  const activeAt = next.activeSub ? findSub(next, next.activeSub) : null;
  if (activeAt) { next.activeGroup = activeAt.gi; next.activeMid = activeAt.mi; }

  const grp = next.groups[next.activeGroup];
  if (next.activeMid < 0 || next.activeMid >= grp.midGroups.length) {
    const mid = newMid(selectedMaktx);
    grp.midGroups.push(mid);
    next.activeMid = grp.midGroups.length - 1;
    next.activeSub = mid.subGroups[0].id;
  }

  const mid = grp.midGroups[next.activeMid];
  if (mid.subGroups.length === 0) mid.subGroups.push(newSub());
  if (!next.activeSub || !mid.subGroups.some((s) => contains(s, next.activeSub!))) next.activeSub = mid.subGroups[0].id;

  const sub = findSub(next, next.activeSub)!.sub;
  if (!sub.lineShifts.includes(lineShift)) sub.lineShifts.push(lineShift);
  return next;
}

// ── Materials list (middle panel) ───────────────────────────────────────────

export function resolveCategory(mtype: string, map: Record<string, string>): string {
  if (!mtype) return "";
  const key = Object.keys(map).find((k) => k.toLowerCase() === mtype.toLowerCase());
  return (key && map[key]) || mtype;
}

/** AND of the whitespace-separated tokens, each found in Maktx, Mtype or the category (any order). */
export function filterMaterials(rows: MaterialRow[], search: string, map: Record<string, string>): MaterialRow[] {
  const tokens = search.trim().toLowerCase().split(/\s+/).filter(Boolean);
  if (tokens.length === 0) return rows;
  return rows.filter((r) => {
    const hay = `${r.maktx}${r.mtype}${resolveCategory(r.mtype, map)}`.toLowerCase();
    return tokens.every((t) => hay.includes(t));
  });
}

/** Materials named by ?m=... (the FCOST unmapped filter) and the ones BMES does not have. */
export function applyFocus(rows: MaterialRow[], focus: string[]): { rows: MaterialRow[]; missing: string[] } {
  const wanted = new Set(focus.map((m) => m.trim().toLowerCase()).filter(Boolean));
  if (wanted.size === 0) return { rows, missing: [] };
  const cmp = (a: string, b: string) => a.localeCompare(b, undefined, { sensitivity: "base" });
  const shown = rows.filter((r) => wanted.has(r.maktx.toLowerCase())).sort((a, b) => cmp(a.maktx, b.maktx));
  const present = new Set(shown.map((r) => r.maktx.toLowerCase()));
  const missing = [...new Set(focus.map((m) => m.trim()).filter((m) => m && !present.has(m.toLowerCase())))].sort(cmp);
  return { rows: shown, missing };
}

export function focusFromQuery(search: string): string[] {
  return [...new Set(new URLSearchParams(search).getAll("m").map((v) => v.trim()).filter(Boolean))];
}

export function saveMessage(c: { groups: number; midGroups: number; subGroups: number; lineShifts: number }): string {
  return `Saved ${c.groups} group(s), ${c.midGroups} mid-group(s), ${c.subGroups} sub-group(s), ${c.lineShifts} line-shift(s).`;
}
