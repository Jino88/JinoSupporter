import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ModelDocs, { DocLinks } from "./ModelDocs";
import { blankRow } from "./processList";
import type { ManualGroup } from "./manual";

describe("Model Info › Documents", () => {
  const html = renderToStaticMarkup(<ModelDocs onSaved={() => {}} onOpen={() => {}} onError={() => {}}
    model={{ name: "MSM-X526", revision: 1, updatedAt: "", updatedBy: "", info: {} } as never} />);

  it("has one AI button that takes several documents at once, then the Work Instructions to read", () => {
    expect(html).toContain("✦ AI: read the documents");
    // one picker for every document, several files at a time; no separate workbook picker
    expect(html.match(/type="file"[^>]*multiple/g)).toHaveLength(1);
    expect(html).not.toContain("Choose workbook");
    expect(html).not.toContain("JIG / Machine / Tester list</h3>");
    expect(html.match(/>Work Instructions</g)).toHaveLength(1);
    expect(html.indexOf("wk-docs-ai")).toBeLessThan(html.indexOf(">Work Instructions<"));
  });

  it("shows how the documents relate: each process with its materials, machine and JIGs, and what is in one only", () => {
    const group = { materials: ["UPPER CASE"], processes: [
      { id: "p1", type: "SUB", name: "Xếp UPPER", localName: "", ngItems: [], materials: [{ material: "UPPER CASE", qtyPerUnit: 1 }] },
      { id: "p2", type: "MAIN", name: "Kiểm tra", localName: "", ngItems: [], materials: [] },
    ] } as unknown as ManualGroup;
    const rows = [
      { ...blankRow("01"), process: "Xếp UPPER", machine: "IONIZER", jigs: [{ name: "CAP JIG", drawingNo: "", ver: "V1", qty: "6" }] },
      { ...blankRow("02"), process: "DRY", machine: "CONVEYOR 3M" },
    ];
    const out = renderToStaticMarkup(<DocLinks result={{ process: { processes: 2, materials: 1, note: "", model: "", saved: true, group },
      jig: { processes: 2, jigKinds: 1, note: "", unmatched: 0, saved: true, rows } }} />);
    expect(out).toContain("1 / 3 processes are in both Process and JIG/MC List");
    expect(out).toMatch(/<b>Xếp UPPER<\/b><\/td><td class="l">UPPER CASE<\/td><td class="l wk-docs-pre">IONIZER<\/td><td class="l wk-docs-pre">CAP JIG<\/td>/);
    expect(out).toContain("<b>DRY</b><span class=\"muted\"> · not in Process</span>");
    expect(out).toContain("<b>Kiểm tra</b><span class=\"muted\"> · not in JIG/MC List</span>");
    expect(renderToStaticMarkup(<DocLinks result={{}} />)).toBe("");
  });
});
