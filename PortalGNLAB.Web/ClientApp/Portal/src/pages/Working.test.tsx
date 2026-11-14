import { describe, expect, it, vi } from "vitest";
import { CATEGORIES, keepDocs, tabOfQuery } from "./Working";

vi.mock("./working/ModelDocs", () => ({ default: () => null }));

describe("Working model page menu", () => {
  const labels = CATEGORIES.map((c) => c.label);
  it("has Model Info, Documents, BOM / Drawing, Process and JIG/MC List all under Model Info", () => {
    expect(labels).toEqual(["Model Info", "Development", "Mass Production", "AI"]);
    expect(CATEGORIES[0].tabs).toEqual([
      ["info", "Model Info"], ["docs", "Documents"], ["bom", "BOM / Drawing"], ["process", "Process"], ["jigmc", "JIG/MC List"],
    ]);
  });
  it("opens the old ?tab=instruction on Documents and keeps the other tabs", () => {
    expect(tabOfQuery("instruction")).toBe("docs");
    expect(tabOfQuery("process")).toBe("process");
    expect(tabOfQuery("jigmc")).toBe("jigmc");
    expect(tabOfQuery("nope")).toBe("info");
    expect(tabOfQuery(null)).toBe("info");
  });
  it("keeps Documents once opened for the model, so its run goes on behind another tab", () => {
    expect(keepDocs("docs", "", "M")).toBe(true);
    expect(keepDocs("jigmc", "M", "M")).toBe(true);
    expect(keepDocs("jigmc", "", "M")).toBe(false);
    expect(keepDocs("jigmc", "M", "N")).toBe(false);
    expect(keepDocs("info", "", null)).toBe(false);
  });
});
