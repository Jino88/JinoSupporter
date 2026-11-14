import { describe, expect, it } from "vitest";
import {
  addMaterial, addMaterials, aiAnswer, blankProcess, bomMaterials, deleteMaterial, emptyManual, moveProcess, ngFromText, processProblem, putMaterial, readManual,
  editProcess, keepSubTables, readSubManual, saveProcess, shownTable, usage, bomGlue, deleteMaterials, glueOf, isBond,
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

  it("takes the BOM's direct materials (depth 1) the list lacks, each once, and adds them", () => {
    const row = (level: number, materialName: string, materialCode = "") => ({ level, materialName, materialCode });
    const g = { materials: ["ENCLOSURE UPPER"], processes: [] };
    const rows = [row(1, "enclosure upper"), row(1, " PLATE 778 "), row(2, "SCREW M2"), row(1, "plate 778"), row(1, "", "Z-CODE-1"), row(3, "DEEP")];
    const names = bomMaterials(g, rows);
    expect(names).toEqual(["PLATE 778", "Z-CODE-1"]);
    expect(addMaterials(g, names).materials).toEqual(["ENCLOSURE UPPER", "PLATE 778", "Z-CODE-1"]);
    expect(addMaterials(g, ["Enclosure Upper", " "])).toBe(g);
  });

  it("puts the materials dragged onto a process at 1 per unit, each once, in list order", () => {
    const keep = proc("AOI 1", { materials: [{ material: "UPPER", qtyPerUnit: 2 }] });
    const other = proc("SPK Assy");
    const g = { materials: ["UPPER", "SPK 201507", "LOWER"], processes: [keep, other] };
    const once = putMaterial(g, other.id, ["spk 201507"]);
    expect(once.processes[1].materials).toEqual([{ material: "SPK 201507", qtyPerUnit: 1 }]);
    expect(once.processes[0]).toBe(keep);
    expect(putMaterial(once, other.id, ["SPK 201507"])).toBe(once);
    expect(putMaterial(g, keep.id, ["upper"])).toBe(g);
    expect(putMaterial(g, other.id, ["NOT LISTED"])).toBe(g);
    // several picked with Ctrl+click: the ones the process lacks, in list order
    expect(putMaterial(g, keep.id, ["LOWER", "UPPER", "SPK 201507"]).processes[0].materials).toEqual([
      { material: "UPPER", qtyPerUnit: 2 }, { material: "SPK 201507", qtyPerUnit: 1 }, { material: "LOWER", qtyPerUnit: 1 },
    ]);
  });

  it("reads the table AI answered: what it added and removed, and materials missing from the list", () => {
    const a = proc("AOI 1"), b = proc("Old step");
    const g = { materials: ["UPPER"], processes: [a, b] };
    const r = aiAnswer(g, {
      materials: ["UPPER"],
      processes: [
        { ...a, ngItems: ["NG", " ng ", ""], materials: [{ material: "LOWER", qtyPerUnit: 2 }] },
        { id: "", type: "ODD", name: " New ", localName: "", ngItems: ["NG"], materials: [] },
      ],
    });
    expect(r.group.materials).toEqual(["UPPER", "LOWER"]);
    expect(r.group.processes[0]).toMatchObject({ id: a.id, name: "AOI 1", ngItems: ["NG"], materials: [{ material: "LOWER", qtyPerUnit: 2 }] });
    expect(r.group.processes[1]).toMatchObject({ type: "SUB", name: "New" });
    expect(r.group.processes[1].id).not.toBe("");
    expect([r.added, r.removed]).toEqual([["New"], ["Old step"]]);
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

  it("puts a process changed in place into the table as typed, the others unchanged", () => {
    const a = proc("AOI 1"), b = proc("SPK Assy");
    const g = { materials: [], processes: [a, b] };
    const next = editProcess(g, { ...b, name: "SPK Assy ", localName: " Tổ hợp" });
    expect(next.processes.map((p) => [p.name, p.localName])).toEqual([["AOI 1", ""], ["SPK Assy ", " Tổ hợp"]]);
    expect(next.processes[0]).toBe(a);
    expect(editProcess(g, proc("Other"))).toEqual(g);
  });

  it("keeps a process table per sub group, read leniently, and drops the table of a removed sub group", () => {
    const top = { materials: ["UPPER"], processes: [proc("Bonding")] };
    const t = readSubManual({ top, bottom: { processes: [{ name: "AOI" }] }, junk: 3 });
    expect(Object.keys(t)).toEqual(["top", "bottom", "junk"]);
    expect(t.top).toEqual(top);
    expect(t.bottom.processes.map((p) => p.name)).toEqual(["AOI"]);
    expect(t.junk).toEqual(emptyManual());
    expect(readSubManual(undefined)).toEqual({});
    expect(readSubManual(["x"])).toEqual({});
    expect(Object.keys(keepSubTables(t, ["bottom", "top"]))).toEqual(["top", "bottom"]);
    expect(keepSubTables({ top, bottom: top }, ["top", "bottom"])).toEqual({ top, bottom: top });
    const same = { top };
    expect(keepSubTables(same, ["top"])).toBe(same);
  });

  it("shows the table of the sub group picked, the model's only while there is no sub group or it has something", () => {
    const filled = { materials: ["MSU-L20S15-07"], processes: [] };
    expect(shownTable(null, [], emptyManual())).toEqual({ at: "", model: true });
    expect(shownTable(null, ["top", "bottom"], emptyManual())).toEqual({ at: "top", model: false });
    expect(shownTable("bottom", ["top", "bottom"], emptyManual())).toEqual({ at: "bottom", model: false });
    expect(shownTable("", ["top", "bottom"], emptyManual())).toEqual({ at: "top", model: false });
    expect(shownTable("gone", ["top", "bottom"], emptyManual()).at).toBe("top");
    expect(shownTable(null, ["top"], filled)).toEqual({ at: "", model: true });
    expect(shownTable("top", ["top"], filled)).toEqual({ at: "top", model: true });
  });

  it("reads a saved group leniently", () => {
    expect(readManual(undefined)).toEqual(emptyManual());
    const g = readManual({ materials: ["A", 3], processes: [{ id: "x", type: "ODD", name: "P", ngItems: ["N", 1], materials: [{ material: "A", qtyPerUnit: "2" }] }] });
    expect(g).toEqual({ materials: ["A"], processes: [{ id: "x", type: "SUB", name: "P", localName: "", ngItems: ["N"], materials: [{ material: "A", qtyPerUnit: 2 }] }] });
  });

  it("tells a bond from a part: the glue the Work Instructions name, a glue name, or a BOM material counted by weight", () => {
    const tool = (name: string, note: string) => ({ no: "", name, note });
    const doc = { processes: [
      { tools: [tool("Thiết bị", "NON CONTACT DISPENSOR DT"), tool("Keo", "SL233LF ( Sử dụng trong vòng 24 giờ)"), tool("Kim", "25G - 6.34mm")],
        parts: [{ no: "1", name: "ENCLOSURE", spec: "ENCLOSURE-X316B-UPPER", qty: "1" }] },
      { tools: [tool("Keo", "G05-0001 ( Sử dụng trong vòng 24 giờ)")], parts: [{ no: "2", name: "Bond", spec: "3777 (PUR)", qty: "" }] },
    ] } as unknown as Parameters<typeof glueOf>[0];
    const glue = glueOf(doc);
    expect([...glue].sort()).toEqual(["3777", "G050001", "SL233LF"]);
    // written another way, or with what it is after it, it is still the bond
    expect(["SL-233LF", "G05-0001", "3777 (PUR)", "Glue 233", "본드 A"].filter((m) => isBond(m, glue))).toHaveLength(5);
    expect(["ENCLOSURE-X316B-UPPER", "SR4060", "MSM-X316B CUSHION-SPL VS", "Bonding tape"].filter((m) => isBond(m, glue))).toEqual([]);
    expect(glueOf(null).size).toBe(0);
    const row = (materialName: string, usageUnit: string, level = 1) => ({ level, materialName, materialCode: "", usageUnit });
    expect(bomGlue([row("PLATE-X316B", "PC"), row("SL233LF", "G"), row("SR4060", " ml "), row("Deep glue", "G", 2)])).toEqual(["SL233LF", "SR4060"]);
  });

  it("takes several materials out of the list and of every process", () => {
    const g = { materials: ["UPPER", "SL233LF", "SR4060"], processes: [proc("Bonding", { materials: [{ material: "UPPER", qtyPerUnit: 1 }, { material: "sl233lf", qtyPerUnit: 1 }] })] };
    const out = deleteMaterials(g, ["SL233LF", "SR4060"]);
    expect(out.materials).toEqual(["UPPER"]);
    expect(out.processes[0].materials).toEqual([{ material: "UPPER", qtyPerUnit: 1 }]);
  });
});
