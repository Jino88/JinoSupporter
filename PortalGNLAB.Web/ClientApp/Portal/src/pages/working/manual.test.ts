import { describe, expect, it } from "vitest";
import {
  addMaterial, blankProcess, deleteMaterial, emptyManual, fromProductionInput, moveProcess, ngFromText, processProblem, readManual,
  saveProcess, usage,
} from "./manual";
import type { ManualProcess } from "./manual";

const proc = (name: string, f: Partial<ManualProcess> = {}): ManualProcess => ({ ...blankProcess(), name, ...f });

describe("model group entered by hand", () => {
  it("reads NG items one per line, without empty lines or repeats", () => {
    expect(ngFromText(" Glue overflow \n\nglue OVERFLOW\r\nMisalign\n")).toEqual(["Glue overflow", "Misalign"]);
  });

  it("keeps a material list without repeats and takes a deleted material out of every process", () => {
    let g = addMaterial(addMaterial(emptyManual(), " UPPER "), "SPK 201507");
    expect(g.materials).toEqual(["UPPER", "SPK 201507"]);
    expect(() => addMaterial(g, "upper")).toThrow(/already/);
    expect(() => addMaterial(g, "  ")).toThrow(/material name/);
    g = saveProcess(g, proc("Bonding Enclosure", { materials: [{ material: "UPPER", qtyPerUnit: 1 }] }));
    g = saveProcess(g, proc("SPK Assy", { materials: [{ material: "SPK 201507", qtyPerUnit: 2 }, { material: "upper", qtyPerUnit: 1 }] }));
    expect([usage(g, "UPPER"), usage(g, "SPK 201507")]).toEqual([2, 1]);
    const after = deleteMaterial(g, "Upper");
    expect([after.materials, after.processes.map((p) => p.materials.map((u) => u.material))]).toEqual([["SPK 201507"], [[], ["SPK 201507"]]]);
  });

  it("checks a process: a name, listed materials once each, qty above 0", () => {
    const mats = ["UPPER"];
    expect(processProblem(proc("AOI 1"), mats)).toBe("");
    expect(processProblem(proc(" "), mats)).toMatch(/name is required/);
    expect(processProblem(proc("A", { materials: [{ material: "LOWER", qtyPerUnit: 1 }] }), mats)).toMatch(/not in the material list/);
    expect(processProblem(proc("A", { materials: [{ material: "UPPER", qtyPerUnit: 0 }] }), mats)).toMatch(/more than 0/);
    expect(processProblem(proc("A", { materials: [{ material: "UPPER", qtyPerUnit: 1 }, { material: "upper", qtyPerUnit: 2 }] }), mats)).toMatch(/twice/);
  });

  it("adds a new process at the end, replaces an edited one in place, and moves processes", () => {
    const a = proc(" AOI 1 ", { ngItems: ["Misalign", "misalign", " "] });
    const b = proc("AOI 2");
    let g = saveProcess(saveProcess(emptyManual(), a), b);
    expect(g.processes.map((p) => [p.name, p.ngItems])).toEqual([["AOI 1", ["Misalign"]], ["AOI 2", []]]);
    g = saveProcess(g, { ...g.processes[0], name: "AOI 1A" });
    expect(g.processes.map((p) => p.name)).toEqual(["AOI 1A", "AOI 2"]);
    expect(moveProcess(g, b.id, -1).processes.map((p) => p.name)).toEqual(["AOI 2", "AOI 1A"]);
    expect(moveProcess(g, b.id, 1)).toBe(g);
  });

  it("copies a Production Input setup", () => {
    const g = fromProductionInput({
      materials: ["UPPER"], stages: [],
      processes: [{ id: "p", name: "Bonding", localName: "Bôi keo", segment: "A", type: "SUB",
        ngItems: [{ id: "n1", name: "NG" }, { id: "n2", name: "Glue overflow" }], materials: [{ material: "UPPER", qtyPerUnit: 1 }] }],
    });
    expect(g.materials).toEqual(["UPPER"]);
    expect(g.processes.map((p) => [p.type, p.name, p.localName, p.ngItems, p.materials])).toEqual([
      ["SUB", "Bonding", "Bôi keo", ["NG", "Glue overflow"], [{ material: "UPPER", qtyPerUnit: 1 }]],
    ]);
  });

  it("reads a saved group leniently", () => {
    expect(readManual(undefined)).toEqual(emptyManual());
    const g = readManual({ materials: ["A", 3], processes: [{ id: "x", type: "ODD", name: "P", ngItems: ["N", 1], materials: [{ material: "A", qtyPerUnit: "2" }] }] });
    expect(g).toEqual({ materials: ["A"], processes: [{ id: "x", type: "SUB", name: "P", localName: "", ngItems: ["N"], materials: [{ material: "A", qtyPerUnit: 2 }] }] });
  });
});
