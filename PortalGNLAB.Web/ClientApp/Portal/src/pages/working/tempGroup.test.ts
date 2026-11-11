import { describe, expect, it } from "vitest";
import {
  addShift, addSub, allShifts, emptyTempGroup, hasTempGroup, readTempGroup, removeShift, removeSub, renameSub, subShiftCount, tempCounts,
} from "./tempGroup";

describe("temporary model group", () => {
  it("builds the Model › Sub Group › Shift Group tree by hand", () => {
    let g = addSub(emptyTempGroup(), null, " SUB ");
    g = addSub(g, null, "MAIN");
    const sub = g.subGroups[0].id;
    g = addSub(g, sub, "SPK Assy");
    const spk = g.subGroups[0].children[0].id;
    g = addShift(g, spk, "MSM-A386B_SUB 1");
    g = addShift(g, sub, "MSM-A386B_SUB 2");
    g = addShift(g, null, "MSM-A386B");
    expect(g.subGroups.map((s) => s.name)).toEqual(["SUB", "MAIN"]);
    expect(g.subGroups[0].children[0]).toMatchObject({ name: "SPK Assy", lineShifts: ["MSM-A386B_SUB 1"] });
    expect(allShifts(g)).toEqual(["MSM-A386B", "MSM-A386B_SUB 2", "MSM-A386B_SUB 1"]);
    expect(subShiftCount(g.subGroups[0])).toBe(2);
    expect(tempCounts(g)).toEqual({ subs: 3, shifts: 3 });
    expect(hasTempGroup(g)).toBe(true);
    expect(hasTempGroup(emptyTempGroup())).toBe(false);
  });

  it("refuses empty names, a sibling sub group twice and a shift group already in the group", () => {
    let g = addSub(emptyTempGroup(), null, "SUB");
    const sub = g.subGroups[0].id;
    expect(() => addSub(g, null, "sub")).toThrow("already a sub group");
    expect(() => addSub(g, null, " ")).toThrow("Type a sub group name");
    g = addSub(g, sub, "SUB"); // same name under another parent is fine
    g = addShift(g, sub, "L1_A");
    expect(() => addShift(g, null, "l1_a")).toThrow("already in this model group");
    expect(() => addShift(g, sub, "")).toThrow("Type a shift group name");
    g = addSub(g, null, "MAIN");
    expect(() => renameSub(g, g.subGroups[1].id, "Sub")).toThrow("already a sub group");
    expect(renameSub(g, g.subGroups[1].id, " Main 2 ").subGroups[1].name).toBe("Main 2");
  });

  it("removes a sub group with everything under it, and a shift group from where it is", () => {
    let g = addSub(emptyTempGroup(), null, "SUB");
    const sub = g.subGroups[0].id;
    g = addSub(g, sub, "Inner");
    g = addShift(g, g.subGroups[0].children[0].id, "L1_A");
    g = addShift(g, null, "L2_B");
    expect(allShifts(removeShift(g, "l1_a"))).toEqual(["L2_B"]);
    const gone = removeSub(g, sub);
    expect(gone.subGroups).toEqual([]);
    expect(allShifts(gone)).toEqual(["L2_B"]);
    expect(g.subGroups[0].children).toHaveLength(1); // the input is not changed
  });

  it("reads a saved group leniently", () => {
    const g = readTempGroup({ group: "MSM", lineShifts: ["A_1", 3, " "], subGroups: [{ id: "s", name: "SUB", lineShifts: ["B_1"], children: [{ name: "In" }] }, null] });
    expect(g.group).toBe("MSM");
    expect(g.lineShifts).toEqual(["A_1"]);
    expect(g.subGroups).toHaveLength(1);
    expect(g.subGroups[0]).toMatchObject({ id: "s", name: "SUB", lineShifts: ["B_1"] });
    expect(g.subGroups[0].children[0].id).toBeTruthy();
    expect(readTempGroup(undefined)).toEqual(emptyTempGroup());
  });
});
