import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import MaterialTab from "./MaterialTab";
import { bomMaterialItems, cleanMaterials, materialGroups, materialProblem, moldCounts, readMaterials, type MaterialItem } from "./materialList";
import { emptyTempGroup } from "./tempGroup";
import type { BmesMaterial, BomRow } from "./bmes";
import type { WorkingModel } from "./types";

const bmes: BmesMaterial = { group: "MSM", productGroup: "MSM", material: "MSM-A386B", products: ["MSM-A386B", "MSM-A386B_SUB"], subGroups: [] };
const item = (id: string, code: string, name: string, mold: MaterialItem["mold"] = ""): MaterialItem => ({ id, code, name, qty: "", unit: "", mold, note: "" });
const bom = (level: number, materialCode: string, materialName: string, usageQty: string | null = "1.000000", usageUnit = "EA"): BomRow =>
  ({ level, materialCode, materialName, usageQty, usageUnit, parentMaterialName: "", hasChildren: false });

describe("Material tab groups", () => {
  it("has a group per product of the BMES model group, else per sub-group of the temporary group, else the model", () => {
    expect(materialGroups("MSM-A386B", bmes, emptyTempGroup(), {}).map((g) => [g.key, g.label, g.product]))
      .toEqual([["p:MSM-A386B", "MSM-A386B", "MSM-A386B"], ["p:MSM-A386B_SUB", "MSM-A386B_SUB", "MSM-A386B_SUB"]]);
    const temp = { group: "TIM", lineShifts: [], subGroups: [{ id: "s1", name: "TOP", lineShifts: [], children: [{ id: "s2", name: "LEFT", lineShifts: [], children: [] }] }] };
    expect(materialGroups("TIM-1", null, temp, {}).map((g) => [g.key, g.label, g.product])).toEqual([["s:s1", "TOP", "TOP"], ["s:s2", "TOP › LEFT", "LEFT"]]);
    expect(materialGroups("TIM-1", null, emptyTempGroup(), {}).map((g) => g.key)).toEqual(["m"]);
  });

  it("keeps showing a saved group that is no longer in BMES, so its materials are not hidden", () => {
    const saved = { "p:OLD": { label: "OLD", rows: [item("a", "C1", "Cover")] }, "p:GONE": { label: "GONE", rows: [] } };
    expect(materialGroups("MSM-A386B", bmes, emptyTempGroup(), saved).map((g) => g.key)).toEqual(["p:MSM-A386B", "p:MSM-A386B_SUB", "p:OLD"]);
  });
});

describe("Material tab rows", () => {
  it("loads the direct materials of the BOM not listed yet (by code), with their usage, or every level when asked", () => {
    const rows = [bom(1, "C1", "Cover"), bom(2, "R1", "Resin"), bom(1, "T1", "Tape", "0.250000", "M"), bom(1, "C1", "Cover again"), bom(1, "", "No code")];
    const got = bomMaterialItems([item("x", "c1", "Cover")], rows);
    expect(got.map((m) => [m.code, m.name, m.qty, m.unit, m.mold])).toEqual([["T1", "Tape", "0.25", "M", ""], ["", "No code", "1", "EA", ""]]);
    expect(bomMaterialItems([], rows, true).map((m) => m.code)).toEqual(["C1", "R1", "T1", ""]);
  });

  it("asks for a code or a name and a code once per group, and counts the molds", () => {
    expect(materialProblem([item("a", "", " ")])).toBe("Row 1 needs a code or a name.");
    expect(materialProblem([item("a", "C1", ""), item("b", "c1 ", "x")])).toBe("The code c1 is listed twice.");
    expect(materialProblem([item("a", "", "Cover"), item("b", "", "Cover")])).toBe("");
    expect(moldCounts([item("a", "1", "", "mass"), item("b", "2", "", "sample"), item("c", "3", "", "sample"), item("d", "4", "")])).toEqual({ mass: 1, sample: 2 });
  });

  it("reads what was saved, a mold other than mass / sample as not set, and saves without empty groups", () => {
    const read = readMaterials({ "p:A": { label: "A", rows: [{ id: "1", code: " C1 ", name: "Cover", qty: 2, mold: "mass" }, { id: "2", code: "C2", mold: "proto" }, null] }, bad: 3 });
    expect(read["p:A"].rows.map((r) => [r.code, r.qty, r.mold])).toEqual([[" C1 ", "2", "mass"], ["C2", "", ""]]);
    expect(readMaterials(null)).toEqual({});
    expect(cleanMaterials({ ...read, "p:B": { label: "B", rows: [] } })).toEqual({ "p:A": { label: "A", rows: [
      { id: "1", code: "C1", name: "Cover", qty: "2", unit: "", mold: "mass", note: "" }, { id: "2", code: "C2", name: "", qty: "", unit: "", mold: "", note: "" },
    ] } });
  });

  it("shows the group chips, the materials with their mold, + Material and Load from BOM", () => {
    const model: WorkingModel = { name: "MSM-A386B", revision: 3, updatedAt: "", updatedBy: "", info: { materials: {
      "p:MSM-A386B": { label: "MSM-A386B", rows: [item("a", "C1", "Cover", "mass"), item("b", "T1", "Tape", "sample"), item("c", "S1", "Screw")] },
    } } };
    const html = renderToStaticMarkup(<MaterialTab model={model} initial={[bmes]} onDirty={() => {}} onSaved={() => {}} onError={() => {}} />);
    expect(html).toContain('aria-selected="true" class="wk-chip on">MSM-A386B<span class="cnt">3</span>');
    expect(html).toContain(">MSM-A386B_SUB</button>");
    expect(html).toContain('<tr class="mold-mass">');
    expect(html).toContain('<tr class="mold-sample">');
    expect(html).toContain('<option value="mass" selected="">Mass production mold</option>');
    expect(html).toContain('<option value="sample" selected="">Sample mold</option>');
    expect(html).toContain("Mass production mold 1</span>");
    expect(html).toContain("Sample mold 1</span>");
    expect(html).toContain("Not set 1</span>");
    expect(html).toContain(">+ Material</button>");
    expect(html).toContain('aria-label="BOM product" value="MSM-A386B"');
  });
});
