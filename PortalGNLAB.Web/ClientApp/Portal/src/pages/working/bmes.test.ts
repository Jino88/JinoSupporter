import { describe, expect, it } from "vitest";
import { drawingsByName, drawingsOf, matchMaterial, qty, visibleBom } from "./bmes";
import type { BmesMaterial, BomRow } from "./bmes";

const mat = (material: string, group = "UNIT"): BmesMaterial => ({ group, productGroup: group, material, products: [], subGroups: [] });
const materials = [mat("TIU-C11-20"), mat("TIU-C11-18"), mat("BRS-161016S08ZZ", "SPK"), mat("MSU-L20S15-07", "SPK")];

describe("BMES model group", () => {
  it("finds the material of a model by its name, without the -ZZ ending", () => {
    expect(matchMaterial("TIU-C11-18-ZZ", materials)).toMatchObject({ material: { material: "TIU-C11-18" }, by: "name" });
    expect(matchMaterial("BRS-161016S08ZZ", materials).material?.material).toBe("BRS-161016S08ZZ");
    expect(matchMaterial("msu-l20s15-07-zz", materials).material?.material).toBe("MSU-L20S15-07");
    expect(matchMaterial("TIU-C11-08", materials)).toEqual({ material: null, by: "" });
  });

  it("uses the material saved in Model Info first", () => {
    expect(matchMaterial("TIU-C11-08", materials, "TIU-C11-20")).toMatchObject({ material: { material: "TIU-C11-20" }, by: "saved" });
    // a saved material that is no longer in BMES falls back to the name
    expect(matchMaterial("TIU-C11-18-ZZ", materials, "GONE")).toMatchObject({ material: { material: "TIU-C11-18" }, by: "name" });
  });

  it("shows BMES quantities without the trailing zeros", () => {
    expect(["1000.000000", "0.250000", "6.296000", 29, 0.25, null, "", "n/a"].map(qty)).toEqual(["1000", "0.25", "6.296", "29", "0.25", "", "", "n/a"]);
  });

  it("hides the rows of a closed branch of the BOM", () => {
    const r = (level: number): BomRow => ({ level, materialCode: "", materialName: "", usageQty: "1", usageUnit: "PC", parentMaterialName: "", hasChildren: false });
    const rows = [r(1), r(2), r(3), r(3), r(2), r(1), r(2)];
    expect(visibleBom(rows, new Set()).map((x) => x.index)).toEqual([0, 1, 2, 3, 4, 5, 6]);
    expect(visibleBom(rows, new Set([1])).map((x) => x.index)).toEqual([0, 1, 4, 5, 6]);
    expect(visibleBom(rows, new Set([0])).map((x) => x.index)).toEqual([0, 5, 6]);
    expect(visibleBom(rows, new Set([0, 5])).map((x) => x.index)).toEqual([0, 5]);
  });

  it("finds the drawings of a material by its name, each PDM number and version once", () => {
    const d = (pdmno: string, pdmvr: string, pdrtx: string) => ({ pdmno, pdmvr, pdtnoTx: "", pdrtx });
    const by = drawingsByName([
      d("P-1", "0", "FRAME-C11-18-L"), d("P-1", "0", "frame-c11-18-l "), d("P-1", "1", "FRAME-C11-18-L"), d("P-2", "0", "COIL-C11-18"), d("P-3", "0", ""),
    ]);
    expect(drawingsOf(by, "Frame-C11-18-L").map((x) => x.pdmno + "/" + x.pdmvr)).toEqual(["P-1/0", "P-1/1"]);
    expect(drawingsOf(by, "COIL-C11-18").length).toBe(1);
    expect(drawingsOf(by, "EA6116R")).toEqual([]);
    expect(by.size).toBe(2);
  });
});
