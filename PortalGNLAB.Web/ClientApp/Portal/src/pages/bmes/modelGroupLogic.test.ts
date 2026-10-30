import { describe, expect, it } from "vitest";
import type { ModelGroup } from "../../bmesApi";
import {
  addGroup, addLineShift, addMid, addSubAtMidTop, addSubUnder, applyFocus, canMoveSub, clearGroups, countSubs, dropSub,
  filterMaterials, focusFromQuery, fromApi, initialEditor, isCleared, midTotal, removeGroup, removeLineShift, removeMid,
  removeSub, renameGroup, resolveCategory, setActiveGroup, setActiveMid, setProductGroup, toApi, type Editor,
} from "./modelGroupLogic";

const saved: ModelGroup[] = [
  {
    name: "G1", productGroup: "SPK",
    midGroups: [{ material: "M1", subGroups: [{ name: "Line 1", lineShifts: ["L1-A"], subGroups: [{ name: "Inner", lineShifts: ["L1-B"], subGroups: [] }] }, { name: "Line 2", lineShifts: [], subGroups: [] }] }],
  },
  { name: "G2", productGroup: "bogus", midGroups: [{ material: "M2", subGroups: [] }] },
];

describe("fromApi / toApi", () => {
  it("round-trips the saved shape, fixing unknown product groups and empty models", () => {
    const e = fromApi(saved);
    expect(e.groups[1].productGroup).toBe("ETC");
    expect(e.groups[1].midGroups[0].subGroups).toHaveLength(1);
    const back = toApi(e.groups);
    expect(back[0]).toEqual(saved[0]);
    expect(back[1]).toEqual({ name: "G2", productGroup: "ETC", midGroups: [{ material: "M2", subGroups: [{ name: "", lineShifts: [], subGroups: [] }] }] });
  });
  it("starts with one empty Group 1 when nothing is saved", () => {
    const e = fromApi([]);
    expect(e.groups.map((g) => g.name)).toEqual(["Group 1"]);
    expect(isCleared(e.groups)).toBe(true);
  });
  it("counts sub groups and line shifts", () => {
    expect(countSubs(saved[0].midGroups[0].subGroups)).toBe(3);
    expect(midTotal(fromApi(saved).groups[0].midGroups[0])).toBe(2);
  });
});

describe("groups", () => {
  it("adds, renames, re-types and removes groups, always keeping one", () => {
    let e = addGroup(initialEditor());
    expect(e.groups.map((g) => g.name)).toEqual(["Group 1", "Group 2"]);
    expect(e.activeGroup).toBe(1);
    e = renameGroup(e, 1, "Speakers");
    e = setProductGroup(e, 1, "UNIT");
    expect(e.groups[1]).toMatchObject({ name: "Speakers", productGroup: "UNIT" });
    e = removeGroup(removeGroup(e, 1), 0);
    expect(e.groups.map((g) => g.name)).toEqual(["Group 1"]);
    expect(clearGroups().groups).toHaveLength(1);
  });
  it("switching group resets the active model and sub group", () => {
    let e = fromApi(saved);
    e = setActiveMid(e, 0, 0);
    expect(e.activeMid).toBe(0);
    e = setActiveGroup(e, 1);
    expect(e).toMatchObject({ activeGroup: 1, activeMid: -1, activeSub: null });
  });
});

describe("line shifts", () => {
  it("creates a model named after the selected material when none is active", () => {
    let e = addLineShift(initialEditor(), "LS-1", "Maktx-9");
    expect(e.groups[0].midGroups[0].material).toBe("Maktx-9");
    expect(e.groups[0].midGroups[0].subGroups[0].lineShifts).toEqual(["LS-1"]);
    e = addLineShift(e, "LS-1", "Maktx-9");
    expect(e.groups[0].midGroups[0].subGroups[0].lineShifts).toEqual(["LS-1"]);
  });
  it("goes into the active sub group, and can be removed again", () => {
    let e = addMid(initialEditor(), 0);
    e = addSubUnder(e, null);
    const sub = e.activeSub!;
    e = addLineShift(e, "LS-2", "");
    expect(e.groups[0].midGroups[0].subGroups[1]).toMatchObject({ id: sub, lineShifts: ["LS-2"] });
    e = removeLineShift(e, sub, "LS-2");
    expect(e.groups[0].midGroups[0].subGroups[1].lineShifts).toEqual([]);
  });
});

describe("sub groups", () => {
  it("names new sub groups Line N and makes them active", () => {
    let e = addMid(initialEditor(), 0);
    e = addSubAtMidTop(e, 0, 0);
    expect(e.groups[0].midGroups[0].subGroups.map((s) => s.name)).toEqual(["", "Line 2"]);
    expect(e.activeSub).toBe(e.groups[0].midGroups[0].subGroups[1].id);
  });
  it("keeps one default sub group when the last one is removed", () => {
    let e = addMid(initialEditor(), 0);
    e = removeSub(e, e.groups[0].midGroups[0].subGroups[0].id);
    expect(e.groups[0].midGroups[0].subGroups).toHaveLength(1);
    expect(e.activeSub).toBeNull();
  });
  it("removes a model and clamps the active model", () => {
    let e = addMid(addMid(initialEditor(), 0), 0);
    e = removeMid(e, 0, 1);
    expect(e.activeMid).toBe(0);
  });
});

describe("moving sub groups", () => {
  const base = (): Editor => fromApi(saved);
  it("allows a move within the same model only, never into itself or a descendant", () => {
    const e = base();
    const [line1, line2] = e.groups[0].midGroups[0].subGroups;
    const inner = line1.subGroups[0];
    expect(canMoveSub(e, line2.id, line1.id)).toBe(true);
    expect(canMoveSub(e, line1.id, line1.id)).toBe(false);
    expect(canMoveSub(e, line1.id, inner.id)).toBe(false);
    expect(canMoveSub(e, line2.id, e.groups[1].midGroups[0].subGroups[0].id)).toBe(false);
  });
  it("drops a sub group into another and reports it", () => {
    const e = base();
    const [line1, line2] = e.groups[0].midGroups[0].subGroups;
    const result = dropSub(e, line2.id, line1.id);
    expect(result.error).toBe(false);
    expect(result.message).toBe("Moved Sub Group 'Line 2' into 'Line 1'.");
    const mid = result.editor.groups[0].midGroups[0];
    expect(mid.subGroups.map((s) => s.name)).toEqual(["Line 1"]);
    expect(mid.subGroups[0].subGroups.map((s) => s.name)).toEqual(["Inner", "Line 2"]);
    expect(result.editor.activeSub).toBe(line2.id);
  });
  it("rejects an invalid drop without changing anything", () => {
    const e = base();
    const line1 = e.groups[0].midGroups[0].subGroups[0];
    const result = dropSub(e, line1.id, line1.subGroups[0].id);
    expect(result.error).toBe(true);
    expect(result.editor).toBe(e);
  });
});

describe("materials list", () => {
  const rows = [{ maktx: "SPK-100 Alpha", mtype: "FERT" }, { maktx: "UNIT-200", mtype: "HALB" }];
  const map = { FERT: "Finished goods" };
  it("resolves the category, case-insensitively, falling back to the type", () => {
    expect(resolveCategory("fert", map)).toBe("Finished goods");
    expect(resolveCategory("HALB", map)).toBe("HALB");
    expect(resolveCategory("", map)).toBe("");
  });
  it("matches every token in any order across maktx, type and category", () => {
    expect(filterMaterials(rows, "finished alpha", map)).toEqual([rows[0]]);
    expect(filterMaterials(rows, "  ", map)).toBe(rows);
    expect(filterMaterials(rows, "zzz", map)).toEqual([]);
  });
  it("applies the ?m= focus filter and lists materials BMES does not have", () => {
    expect(focusFromQuery("?m=UNIT-200&m=Nope&m=%20&m=UNIT-200")).toEqual(["UNIT-200", "Nope"]);
    const r = applyFocus(rows, ["unit-200", "Nope"]);
    expect(r.rows).toEqual([rows[1]]);
    expect(r.missing).toEqual(["Nope"]);
    expect(applyFocus(rows, [])).toEqual({ rows, missing: [] });
  });
});
