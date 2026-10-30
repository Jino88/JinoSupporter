import { describe, expect, it } from "vitest";
import type { BomResult, BomRow } from "../../bmesReportApi";
import {
  buildStandaloneHtml, byProduct, catalogMatchScore, filterCachedModels, findPdm, formatUnitPrice, formatVnd, hasUsefulModelMatch,
  isNoiseCode, levenshtein, localSuggestions, matchedProducts, modelMatchScore, normalizeModelSearch, pdmIndex, sanitize, visibleRows,
} from "./test5Logic";

const row = (o: Partial<BomRow>): BomRow => ({
  productCode: "P-1", productName: "PROD-L", materialCode: "X", materialName: "", bomLevel: 1,
  bomPath: "P-1>X", bomPathText: "PROD-L > X", hasChildren: false, ...o,
});

describe("model search", () => {
  it("normalizes to upper-case letters and digits", () => {
    expect(normalizeModelSearch("ab-12 c_d")).toBe("AB12CD");
  });
  it("scores exact, prefix, contains and fuzzy matches", () => {
    expect(modelMatchScore("abc", "ABC")).toBe(0);
    expect(modelMatchScore("ab", "ABCD")).toBe(12);
    expect(modelMatchScore("bc", "ABCD")).toBe(33);
    expect(modelMatchScore("xyz", "ABC")).toBe(103);
    expect(levenshtein("KITTEN", "SITTING")).toBe(3);
  });
  it("needs a containing match or a 2+ character token", () => {
    expect(hasUsefulModelMatch("X526", "SPK-X526-L")).toBe(true);
    expect(hasUsefulModelMatch("zz 52", "AB52")).toBe(true);
    expect(hasUsefulModelMatch("q", "AB52")).toBe(false);
    expect(catalogMatchScore("q", "AB52")).toBe(2147483647);
    expect(catalogMatchScore("X526", "X526")).toBe(-250);
  });
  it("merges catalog and cached suggestions without duplicates", () => {
    const s = localSuggestions(
      "x526",
      [{ productCode: "P-S-1", productName: "X526-L" }, { productCode: "P-S-2", productName: "" }],
      [{ name: "x526-l", groups: "G" }, { name: "X526-R", groups: "G" }],
    );
    expect(s.map((x) => [x.name, x.source])).toEqual([["X526-L", "catalog"], ["X526-R", "cache"]]);
    expect(localSuggestions(" ", [], [])).toEqual([]);
  });
  it("filters cached models by containment, case-insensitively", () => {
    expect(filterCachedModels([{ name: "ABC", groups: "" }, { name: "XYZ", groups: "" }], " b ").map((m) => m.name)).toEqual(["ABC"]);
  });
});

describe("visibleRows", () => {
  it("hides M-P / C-S lines and pulls their children up", () => {
    const rows = [
      row({ materialCode: "C-S-1", bomLevel: 1, bomPath: "P-1>C-S-1", bomPathText: "PROD-L > ASSY", hasChildren: true }),
      row({ materialCode: "R-1", materialName: "Ring", bomLevel: 2, bomPath: "P-1>C-S-1>R-1", bomPathText: "PROD-L > ASSY > Ring" }),
      row({ materialCode: "R-1", materialName: "Ring", bomLevel: 1, bomPath: "P-1>R-1", bomPathText: "PROD-L > Ring" }),
      row({ materialCode: "m-p-9", bomLevel: 1 }),
    ];
    expect(isNoiseCode(" m-p-9")).toBe(true);
    const v = visibleRows(rows, true);
    expect(v).toHaveLength(1);
    expect(v[0]).toMatchObject({ materialCode: "R-1", bomLevel: 1, bomPathText: "PROD-L > Ring", hasChildren: false });
    expect(visibleRows(rows, false)).toBe(rows);
  });
});

describe("byProduct", () => {
  it("keeps one tree per product without hiding", () => {
    const sections = byProduct([row({ productName: "B" }), row({ productName: "a", bomPathText: "a > X" })], false);
    expect(sections.map((s) => s.product)).toEqual(["a", "B"]);
  });
  it("splits common and side-specific materials when hiding with several products", () => {
    const rows = [
      row({ productName: "L", materialCode: "C1", materialName: "Common" }),
      row({ productName: "R", materialCode: "C1", materialName: "Common", bomLevel: 3 }),
      row({ productName: "L", materialCode: "L1", materialName: "Left" }),
    ];
    const sections = byProduct(rows, true);
    expect(sections.map((s) => [s.product, s.isShared, s.rows.map((r) => r.materialCode)])).toEqual([
      ["Common — all 2 products", true, ["C1"]],
      ["L only", false, ["L1"]],
    ]);
    expect(sections[0].rows[0].bomLevel).toBe(1);
  });
});

describe("formats and helpers", () => {
  it("formats prices with significant digits", () => {
    expect(formatUnitPrice(null)).toBe("");
    expect(formatUnitPrice(0)).toBe("");
    expect(formatUnitPrice(12345.5)).toBe("12,346");
    expect(formatUnitPrice(12.345)).toBe("12.35");
    expect(formatUnitPrice(0.0123456789)).toBe("0.012346");
    expect(formatVnd(1234.5)).toBe("1,235");
  });
  it("sanitizes download names", () => {
    expect(sanitize("--X526 L/R--")).toBe("X526-L-R");
  });
  it("finds PDM rows by material name, first one wins", () => {
    const idx = pdmIndex([
      { pdmno: "D1", pdmvr: "0", pdtnoTx: "T", pdrtx: " Ring " },
      { pdmno: "D2", pdmvr: "0", pdtnoTx: "T", pdrtx: "ring" },
    ]);
    expect(findPdm(idx, row({ materialName: "RING" }))?.pdmno).toBe("D1");
  });
  it("lists matched products and writes the HTML download", () => {
    const result: BomResult = {
      model: "X<1>", rows: [row({ materialName: "Ring & Co" })], error: "", warning: "", loadSource: "cache", fetchedAt: "",
    };
    expect(matchedProducts(result)).toEqual(["PROD-L"]);
    const html = buildStandaloneHtml([result], new Map(), { plant: "3200", maxDepth: 6, workDate: "2026-10-30", hideNoise: true, generated: "2026-10-30 09:00" });
    expect(html).toContain("<title>BOM — X&lt;1&gt;</title>");
    expect(html).toContain("Ring &amp; Co");
    expect(html).toContain("M-P / C-S hidden");
  });
});
