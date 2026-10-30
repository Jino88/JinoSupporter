import { describe, expect, it } from "vitest";
import type { RoutingRow } from "../../bmesApi";
import { filterModels, modelCounts, pickModel, selectedRows, selectedTypes } from "./routingLogic";

const row = (id: number, modelName: string, processCode: string, processName: string, processType: string): RoutingRow =>
  ({ id, modelName, processCode, processName, processType });
const rows = [
  row(1, "B-Model", "P10", "Assembly", "Main"),
  row(2, "A-Model", "P20", "Test", "Main"),
  row(3, "A-Model", "P30", "Pack", "Sub"),
  row(4, "", "P99", "Orphan", ""),
];

describe("modelCounts", () => {
  it("counts rows per model, sorted, skipping empty names", () => {
    expect(modelCounts(rows)).toEqual([{ name: "A-Model", count: 2 }, { name: "B-Model", count: 1 }]);
  });
});

describe("filterModels", () => {
  const models = modelCounts(rows);
  it("returns all models for a blank search", () => expect(filterModels(models, rows, "  ")).toBe(models));
  it("matches the model name", () => expect(filterModels(models, rows, "b-mo").map((m) => m.name)).toEqual(["B-Model"]));
  it("matches a process code or name of the model", () => {
    expect(filterModels(models, rows, "p30").map((m) => m.name)).toEqual(["A-Model"]);
    expect(filterModels(models, rows, "assem").map((m) => m.name)).toEqual(["B-Model"]);
  });
});

describe("selectedRows / selectedTypes", () => {
  it("filters by model and type", () => {
    expect(selectedRows(rows, "A-Model", "").map((r) => r.id)).toEqual([2, 3]);
    expect(selectedRows(rows, "A-Model", "Sub").map((r) => r.id)).toEqual([3]);
    expect(selectedRows(rows, "", "")).toEqual([]);
  });
  it("lists the distinct types of the model", () => expect(selectedTypes(rows, "A-Model")).toEqual(["Main", "Sub"]));
});

describe("pickModel", () => {
  it("keeps the current model while it has rows", () => expect(pickModel("B-Model", rows)).toBe("B-Model"));
  it("falls back to the first model", () => expect(pickModel("Gone", rows)).toBe("A-Model"));
  it("is empty without rows", () => expect(pickModel("", [])).toBe(""));
});
