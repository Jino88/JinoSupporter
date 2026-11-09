import { describe, expect, it } from "vitest";
import { buildTree } from "./logic";
import { ADDED_GROUP, allModels, groupChoices, hideModels, withAdded } from "./tree";
import type { ModelGroup } from "./types";

const g = (id: string, name: string, parentId: string | null, models: string[], sort = 0): ModelGroup => ({ id, name, parentId, models, sort, position: 0 });
const groups = [g("spk", "SPEAKER", null, []), g("msu", "MSU", "spk", ["MSU-L15S10-20-ZZ", "MSU-L20S15-07-ZZ"]), g("unit", "UNIT", null, ["TIU-C11-08", "TIU-C11-18-ZZ"], 1)];
const added = (name: string, groupId: string) => ({ name, groupId, createdBy: "jh", createdAt: "2026-11-09T19:00:00Z" });

describe("model tree settings", () => {
  it("puts added models into their group, and the ones without a known group into Added at the end", () => {
    const tree = buildTree(withAdded(groups, [added("MSU-NEW-01", "msu"), added("ODD-1", ""), added("ODD-2", "gone")]));
    expect(allModels(tree)).toEqual(["MSU-L15S10-20-ZZ", "MSU-L20S15-07-ZZ", "MSU-NEW-01", "TIU-C11-08", "TIU-C11-18-ZZ", "ODD-1", "ODD-2"]);
    expect(tree.map((n) => [n.name, n.total])).toEqual([["SPEAKER", 3], ["UNIT", 2], ["Added", 2]]);
    expect(withAdded(groups, []).some((x) => x.id === ADDED_GROUP)).toBe(false);
  });

  it("hides the chosen models, drops groups left empty and recounts", () => {
    const tree = buildTree(groups);
    const shown = hideModels(tree, ["msu-l15s10-20-zz", " MSU-L20S15-07-ZZ", "TIU-C11-08"]);
    expect(shown.map((n) => [n.name, n.total, n.models])).toEqual([["UNIT", 1, ["TIU-C11-18-ZZ"]]]);
    expect(hideModels(tree, [])).toBe(tree);
  });

  it("lists the groups a model can be added to, with their path", () => {
    expect(groupChoices(buildTree(withAdded(groups, [added("X", "")])))).toEqual([["spk", "SPEAKER"], ["msu", "SPEAKER › MSU"], ["unit", "UNIT"]]);
  });
});
