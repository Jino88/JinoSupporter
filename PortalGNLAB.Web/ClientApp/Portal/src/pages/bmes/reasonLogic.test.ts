import { describe, expect, it } from "vitest";
import type { ReasonRow } from "../../bmesApi";
import { MISSING_KEY, chips, displayRows, findPairsMessage, matchesChip, quickOptions, reasonOptions } from "./reasonLogic";

const row = (id: number, processName: string, ngName: string, reason: string): ReasonRow => ({ id, processName, ngName, reason });
const rows = [
  row(1, "Weld", "Crack", "Material"),
  row(2, "Weld", "Burn", " material "),
  row(3, "Pack", "Dent", ""),
  row(4, "Pack", "Scratch", "Handling"),
];

describe("reasonOptions", () => {
  it("is distinct, trimmed, case-insensitive and sorted", () => expect(reasonOptions(rows)).toEqual(["Handling", "Material"]));
});

describe("chips", () => {
  it("lists All, Missing and reasons by count", () => {
    expect(chips(rows)).toEqual([
      { key: "", label: "All", count: 4 },
      { key: MISSING_KEY, label: "Missing reason", count: 1 },
      { key: "Material", label: "Material", count: 2 },
      { key: "Handling", label: "Handling", count: 1 },
    ]);
  });
});

describe("matchesChip", () => {
  it("handles all, missing and a reason", () => {
    expect(matchesChip("", "x")).toBe(true);
    expect(matchesChip(MISSING_KEY, "  ")).toBe(true);
    expect(matchesChip(MISSING_KEY, "x")).toBe(false);
    expect(matchesChip("Material", " MATERIAL ")).toBe(true);
  });
});

describe("displayRows", () => {
  const base = { rows, chip: "", search: "", editId: -1, modelPairs: null };
  it("filters the table by chip and search", () => {
    expect(displayRows({ ...base, chip: "Material" }).map((d) => d.row?.id)).toEqual([1, 2]);
    expect(displayRows({ ...base, search: "dent" }).map((d) => d.row?.id)).toEqual([3]);
  });
  it("keeps the row being edited first even when filtered out", () => {
    expect(displayRows({ ...base, chip: "Material", editId: 4 }).map((d) => d.row?.id)).toEqual([4, 1, 2]);
  });
  it("marks pairs that are not in the table", () => {
    const view = displayRows({ ...base, modelPairs: [{ processName: "Weld", ngName: "Crack" }, { processName: "Weld", ngName: "New" }] });
    expect(view.map((d) => d.row?.id ?? null)).toEqual([1, null]);
    expect(displayRows({ ...base, chip: "Handling", modelPairs: [{ processName: "Weld", ngName: "New" }] })).toEqual([]);
  });
});

describe("quickOptions", () => {
  it("adds a current value that is not a known reason", () => {
    expect(quickOptions(["A"], "b")).toEqual(["b", "A"]);
    expect(quickOptions(["A"], "a")).toEqual(["A"]);
  });
});

describe("findPairsMessage", () => {
  it("reports additions and failures", () => {
    expect(findPairsMessage({ apiAdded: 2, cacheAdded: 1, apiFailure: null, cacheFailure: null }))
      .toEqual({ text: "Added 3 new pair(s) (API 2, cache 1). Click Update To Server to sync.", error: false, added: true });
    const none = findPairsMessage({ apiAdded: 0, cacheAdded: 0, apiFailure: "BMES API step failed", cacheFailure: null });
    expect(none.error).toBe(true);
    expect(none.text).toBe("No new pairs. The Reason Table already covers all known values. BMES API step failed.");
  });
});
