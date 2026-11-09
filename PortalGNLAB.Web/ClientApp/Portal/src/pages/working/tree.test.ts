import { describe, expect, it } from "vitest";
import { buildTree } from "./logic";
import { ADDED_GROUP, BMES_GROUP, allModels, groupChoices, hideModels, withAdded, withBmes } from "./tree";
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

  it("shows sub-groups made in Manage under their group, with the models moved and added into them", () => {
    const sub = (id: string, name: string, parentId: string, sort = 0) => ({ id, name, parentId, sort });
    const tree = buildTree(withAdded(
      groups,
      [added("MSU-NEW-01", "ug-l20"), added("ODD-1", "ug-gone")],
      [sub("ug-l20", "L20 series", "msu"), sub("ug-deep", "Deep", "ug-l20"), sub("ug-unit", "Old units", "unit", 1)],
      [{ name: "msu-l20s15-07-zz", groupId: "ug-l20" }, { name: "TIU-C11-08", groupId: "ug-unit" }, { name: "TIU-C11-18-ZZ", groupId: "ug-gone" }],
    ));
    const msu = tree[0].children[0];
    expect([msu.models, msu.children.map((c) => [c.name, c.models, c.children.map((d) => d.name)])]).toEqual([
      ["MSU-L15S10-20-ZZ"], [["L20 series", ["msu-l20s15-07-zz", "MSU-NEW-01"], ["Deep"]]],
    ]);
    // a move to a group that is gone leaves the model where it was; an added model with no known group goes to Added
    expect(tree.map((n) => [n.name, n.models, n.children.map((c) => [c.name, c.models])])).toEqual([
      ["SPEAKER", [], [["MSU", ["MSU-L15S10-20-ZZ"]]]],
      ["UNIT", ["TIU-C11-18-ZZ"], [["Old units", ["TIU-C11-08"]]]],
      ["Added", ["ODD-1"], []],
    ]);
  });

  it("puts the BMES Make Model Group models in as they are, into the group of the same name, without repeating a model", () => {
    const gloss = [g("spk", "SPEAKER", null, []), g("msu", "MSU", "spk", ["MSU-L20S15-07-ZZ"]), g("tws", "TWS", null, ["TIM-338"], 1),
      g("unit", "UNIT", null, ["TIU-C11-20-ZZ"], 2)];
    const bmes = [
      { group: "TWS", material: "ASSY 338 RA1" }, { group: "TWS", material: "ASSY REAR-TAPE-338" },
      { group: "SPK", material: "MSU-L20S15-07" }, { group: "SPK", material: "BRS-NEW-01" },
      { group: "UNIT", material: "tiu-c11-20" }, { group: "MODULE", material: "MSM-X526/X626B" }, { group: "MODULE", material: "MSM-NX14" },
      { group: "ETC", material: "JIG-1" }, { group: "ETC", material: " " },
    ];
    const r = withBmes(gloss, bmes, [added("MSM-X526/X626B", "")]);
    expect(r.bmesModels).toEqual(["ASSY 338 RA1", "ASSY REAR-TAPE-338", "BRS-NEW-01", "MSM-NX14", "JIG-1"]);
    const tree = buildTree(r.groups);
    expect(tree.map((n) => [n.name, n.models])).toEqual([
      ["SPEAKER", ["BRS-NEW-01"]], ["TWS", ["TIM-338", "ASSY 338 RA1", "ASSY REAR-TAPE-338"]], ["UNIT", ["TIU-C11-20-ZZ"]],
      ["MODULE", ["MSM-NX14"]], ["ETC", ["JIG-1"]],
    ]);
    expect(r.groups.find((x) => x.name === "MODULE")!.id).toBe(BMES_GROUP + "module");
    // the Glossary groups themselves are not changed
    expect(gloss[2].models).toEqual(["TIM-338"]);
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
