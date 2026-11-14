import { describe, expect, it, vi } from "vitest";
import { CATEGORIES, tabOfQuery } from "./Working";

vi.mock("./working/ModelDocs", () => ({ default: () => null }));

describe("Working model page menu", () => {
  const labels = CATEGORIES.map((c) => c.label);
  it("has Model Info with Documents, then Process and JIG/MC List, and no Process List or Work Instructions", () => {
    expect(labels).toEqual(["Model Info", "BOM / Drawing", "Development", "Mass Production", "Process", "JIG/MC List", "AI"]);
    expect(CATEGORIES[0].tabs).toEqual([["info", "Model Info"], ["docs", "Documents"]]);
    expect(CATEGORIES.find((c) => c.label === "Process")!.tabs).toEqual([["process", "Process"]]);
    expect(CATEGORIES.find((c) => c.label === "JIG/MC List")!.tabs).toEqual([["jigmc", "JIG/MC List"]]);
  });
  it("opens the old ?tab=instruction on Documents and keeps the other tabs", () => {
    expect(tabOfQuery("instruction")).toBe("docs");
    expect(tabOfQuery("process")).toBe("process");
    expect(tabOfQuery("jigmc")).toBe("jigmc");
    expect(tabOfQuery("nope")).toBe("info");
    expect(tabOfQuery(null)).toBe("info");
  });
});
