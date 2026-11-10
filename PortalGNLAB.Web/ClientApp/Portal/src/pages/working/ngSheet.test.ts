import { describe, expect, it } from "vitest";
import { blankProcess } from "./manual";
import { mergeSheet, parseNgHeader } from "./ngSheet";

const RATE = "NG rate (%)";
const D3 = ["NG tràn keo", "NG thiếu keo", "NG Lệch"];

// Part of the Daily NG Rate header of 2026-11-10 (a "" item is an empty merged column, rate=false has no NG rate column).
const SHEET: [string, string[], boolean][] = [
  ["1. BONDING ENCLOSURE \nBôi keo tổ hợp spk", ["NG"], true],
  ["2. AOI 1", D3, true],
  ["4. FIX BONDING  SPK 1\n Bôi keo cố định spk 1", ["NG"], true],
  ["15. SPOT WELDING\nHàn", ["Hàn non \nWeak solder)"], true],
  [" 21.  LOWER ASSY\nTổ hợp lower ", ["NG"], true],
  [" 23. VISION INSPECTION \nKiểm tra keo", D3, true],
  ["27.VISION INSPECTION \nKiểm tra sau dán", ["NG", ""], true],
  ["32. BAKO TEST", ["THD", "THD + FRF", "RB"], true],
  ["35. Kiểm tra ANTENA", ["Xước ANTENA", "Gãy ANTENA", "Other"], false],
];

/** The header rows as Excel copies them: merged cells as empty cells, tab separated, line breaks quoted. */
function sheetText(sheet = SHEET, lead = 0): string {
  const head: string[] = Array(lead).fill("");
  const items: string[] = Array(lead).fill("");
  const ba: string[] = Array(lead).fill("");
  if (lead) head[0] = "Date";
  for (const [cell, ngs, rate] of sheet) {
    const cols = ["Input", ...ngs.flatMap((n) => [n, ""]), ...(rate ? [RATE, ""] : [])];
    head.push(cell, ...Array(cols.length - 1).fill(""));
    items.push(...cols);
    ba.push("", ...Array((cols.length - 1) / 2).fill(["Before", "After"]).flat());
  }
  const q = (c: string) => (/[\t\n"]/.test(c) ? `"${c.replace(/"/g, '""')}"` : c);
  return [head, items, ba].map((r) => r.map(q).join("\t")).join("\r\n") + "\r\n";
}

describe("Daily NG Rate header paste", () => {
  it("reads the processes, local names and NG items", () => {
    const { processes: p, warnings, problem } = parseNgHeader(sheetText(SHEET, 4));
    expect(problem).toBe("");
    expect(p).toHaveLength(9);
    expect(p[0]).toEqual({ name: "BONDING ENCLOSURE", localName: "Bôi keo tổ hợp spk", ngItems: ["NG"] });
    expect(p[1]).toEqual({ name: "AOI 1", localName: "", ngItems: D3 });
    expect(p[2].name).toBe("FIX BONDING SPK 1");
    expect(p[3].ngItems).toEqual(["Hàn non Weak solder)"]);
    expect(p[4].name).toBe("LOWER ASSY");
    expect(p[6]).toEqual({ name: "VISION INSPECTION", localName: "Kiểm tra sau dán", ngItems: ["NG"] });
    expect(p[8].ngItems).toEqual(["Xước ANTENA", "Gãy ANTENA", "Other"]);
    expect(warnings).toEqual(["VISION INSPECTION appears 2 times."]);
  });

  it("takes the columns under a process name when it has no Input cell", () => {
    const rows = sheetText([["1. A", ["x"], true], ["2. B", ["y", "z"], true]]).split("\r\n").map((r) => r.split("\t"));
    rows[1][rows[1].indexOf("Input", 1)] = "";
    const r = parseNgHeader(rows.map((x) => x.join("\t")).join("\r\n"));
    expect(r.processes.map((p) => p.ngItems)).toEqual([["x"], ["y", "z"]]);
    expect(r.warnings[0]).toMatch(/2 processes but 1 Input cells/);
  });

  it("says what is missing without an Input row", () => {
    expect(parseNgHeader("1. A\t2. B\r\n").problem).toMatch(/No row with Input/);
  });

  it("keeps the materials and the type and materials of a process with the same name", () => {
    const keep = { ...blankProcess("MAIN"), id: "keep1", name: "aoi 1", localName: "old", ngItems: ["x"], materials: [{ material: "UPPER", qtyPerUnit: 1 }] };
    const gone = { ...blankProcess(), name: "Old step" };
    const { group, dropped } = mergeSheet({ materials: ["UPPER"], processes: [keep, gone] },
      [{ name: "AOI 1", localName: "", ngItems: ["NG"] }, { name: "New", localName: "Mới", ngItems: [] }]);
    expect(group.materials).toEqual(["UPPER"]);
    expect(group.processes[0]).toEqual({ ...keep, name: "AOI 1", localName: "old", ngItems: ["NG"] });
    expect(group.processes[1]).toMatchObject({ type: "SUB", name: "New", localName: "Mới", ngItems: [], materials: [] });
    expect(group.processes[1].id).not.toBe("keep1");
    expect(dropped).toEqual(["Old step"]);
  });
});
