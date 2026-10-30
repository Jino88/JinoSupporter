import { describe, expect, it } from "vitest";
import type { FCostSubGroup, FCostView, PeriodKind } from "../../bmesReportApi";
import { csvText, n0, roundHalfEven } from "./format";
import {
  allColumnIndexes, buildCsvRows, buildHierarchy, colBoundaryClass, computeTotals, defaultStartDate, formatNum, formatRate,
  isRowVisible, parseSubPathSegments, unmappedModelGroupUrl, visibleColumnIndexes, visiblePeriodColumnCount,
} from "./fcostLogic";

const kinds = (...k: PeriodKind[]) => k.map((kind, i) => ({ header: `${kind[0]}${i}`, kind }));

const view = (o: Partial<FCostView> = {}): FCostView => {
  const columns = o.columns ?? kinds("Day", "Day", "Week", "Month");
  return {
    queryDate: "2026-10-30", totalRows: 3, startDate: "2026-08-01", endDate: "2026-10-30",
    headers: columns.map((c) => c.header), columns, total: null, materials: [], subGroups: [], unmapped: [], rawStatus: null,
    ...o,
  };
};

const sg = (o: Partial<FCostSubGroup>): FCostSubGroup => ({
  display: "", groupName: "G", midGroupName: "M", subGroupKey: "G::M::A", matchedMaterialCount: 1,
  inputByCol: [0, 0, 0, 0], costByCol: [0, 0, 0, 0], sourceRateByCol: [0, 0, 0, 0], ...o,
});

describe("number formats", () => {
  it("rounds ties to even like Math.Round and groups thousands", () => {
    expect(roundHalfEven(2.5)).toBe(2);
    expect(roundHalfEven(3.5)).toBe(4);
    expect(roundHalfEven(-2.5)).toBe(-2);
    expect(roundHalfEven(2.6)).toBe(3);
    expect(n0(1234567.4)).toBe("1,234,567");
    expect(n0(-0.4)).toBe("0");
  });
  it("shows a dash for zero", () => {
    expect(formatNum(0)).toBe("-");
    expect(formatNum(1500)).toBe("1,500");
    expect(formatRate(0)).toBe("-");
    expect(formatRate(1.005)).toBe("1.00%");
    expect(formatRate(12.345678)).toBe("12.35%");
  });
});

describe("defaultStartDate", () => {
  it("is the first of the month two months back", () => {
    const d = defaultStartDate(new Date(2026, 0, 15));
    expect([d.getFullYear(), d.getMonth(), d.getDate()]).toEqual([2025, 10, 1]);
  });
});

describe("column folding", () => {
  const many = view({ columns: kinds(...Array<PeriodKind>(10).fill("Day"), "Week", "Week", "Month") });
  it("keeps the first 8 day columns until expanded", () => {
    const folded = visibleColumnIndexes(many, { Day: false, Week: false, Month: false });
    expect(folded).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 10, 11, 12]);
    expect(visiblePeriodColumnCount(many, "Day", { Day: false, Week: false, Month: false })).toBe(8);
    expect(visibleColumnIndexes(many, { Day: true, Week: false, Month: false })).toEqual(allColumnIndexes(many));
  });
  it("uses every header when there is no column metadata", () => {
    expect(visibleColumnIndexes(view({ columns: [], headers: ["C1", "C2"] }), { Day: false, Week: false, Month: false })).toEqual([0, 1]);
  });
  it("marks where the period kind changes", () => {
    const v = view();
    expect([0, 1, 2, 3].map((i) => colBoundaryClass(v, i))).toEqual(["", "", "col-week-start", "col-month-start"]);
  });
});

describe("computeTotals", () => {
  const material = (input: number[], cost: number[]) => ({ displayName: "x", productGroup: "", material: "x", input, cost });
  it("prefers the BMES total when it has input and cost", () => {
    const v = view({ total: material([10, 20, 0, 0], [1, 2, 0, 0]), materials: [material([5, 5, 5, 5], [1, 1, 1, 1])] });
    expect(computeTotals(v, [0, 1])).toEqual({ input: [10, 20], cost: [1, 2] });
  });
  it("falls back to materials, then sub groups, for the selected columns", () => {
    const v = view({ total: material([10, 0, 0, 0], [0, 0, 0, 0]), materials: [material([5, 6, 0, 0], [1, 1, 0, 0])] });
    expect(computeTotals(v, [0, 1])).toEqual({ input: [5, 6], cost: [1, 1] });
    const onlySubs = view({ subGroups: [sg({ inputByCol: [4, 0, 0, 0], costByCol: [2, 0, 0, 0] })] });
    expect(computeTotals(onlySubs, [0])).toEqual({ input: [4], cost: [2] });
  });
});

describe("hierarchy", () => {
  it("parses the path of a sub-group key", () => {
    expect(parseSubPathSegments("G::M::A / B")).toEqual(["A", "B"]);
    expect(parseSubPathSegments("G::M")).toEqual([]);
    expect(parseSubPathSegments("G::M::")).toEqual([]);
  });
  it("builds Mid → Sub1 → Sub2 with summed parents, ordered by group then mid", () => {
    const v = view({
      subGroups: [
        sg({ groupName: "Z", midGroupName: "M2", subGroupKey: "Z::M2::A", display: "A", inputByCol: [1, 0, 0, 0] }),
        sg({ groupName: "G", midGroupName: "M1", subGroupKey: "G::M1::L / L1", display: "L1", inputByCol: [2, 0, 0, 0], costByCol: [1, 0, 0, 0] }),
        sg({ groupName: "G", midGroupName: "M1", subGroupKey: "G::M1::L / L2", display: "L2", inputByCol: [3, 0, 0, 0] }),
      ],
    });
    const rows = buildHierarchy(v);
    expect(rows.map((r) => [r.rowId, r.depth])).toEqual([
      ["M::G::M1", 0],
      ["S1::M::G::M1::L", 1],
      ["S2::G::M1::L / L1", 2],
      ["S2::G::M1::L / L2", 2],
      ["M::Z::M2", 0],
      ["S1::M::Z::M2::A", 1],
    ]);
    expect(rows[0].inputByCol[0]).toBe(5);
    expect(rows[0].costByCol[0]).toBe(1);

    const byId = new Map(rows.map((r) => [r.rowId, r]));
    expect(rows.filter((r) => isRowVisible(r, byId, new Set())).map((r) => r.rowId)).toEqual(["M::G::M1", "M::Z::M2"]);
    expect(isRowVisible(rows[2], byId, new Set(["S1::M::G::M1::L"]))).toBe(false);
    expect(isRowVisible(rows[2], byId, new Set(["M::G::M1", "S1::M::G::M1::L"]))).toBe(true);
  });
});

describe("unmappedModelGroupUrl", () => {
  it("lists each material once, sorted, as m= parameters", () => {
    expect(unmappedModelGroupUrl([{ material: "b" }, { material: "A 1" }, { material: "B" }, { material: "" }]))
      .toBe("/bmes/make-model-group?source=fcost&m=A%201&m=b");
    expect(unmappedModelGroupUrl([])).toBe("/bmes/make-model-group");
  });
});

describe("CSV", () => {
  it("writes summary and material sections with the chosen metrics", () => {
    const v = view({
      columns: kinds("Day", "Week"),
      materials: [{ displayName: "Mat, A", productGroup: "SPK", material: "A", input: [1000, 0], cost: [10, 0] }],
      subGroups: [sg({ inputByCol: [1000, 0], costByCol: [10, 0], subGroupKey: "G::M::A", display: "A" })],
    });
    const rows = buildCsvRows(v, { input: true, cost: false, rate: true });
    expect(rows[0]).toEqual(["F-Cost Report", "2026-08-01 - 2026-10-30"]);
    expect(rows[2]).toEqual(["Section", "Level", "Name", "Type", "D0", "W1"]);
    expect(rows[3]).toEqual(["Model / Sub Group Summary", "Total", "Total", "INPUT", "1,000", "-"]);
    expect(rows[4]).toEqual(["Model / Sub Group Summary", "Total", "Total", "RATE %", "1.00%", "-"]);
    expect(rows.some((r) => r[1] === "Model" && r[2] === "M")).toBe(true);
    expect(rows.at(-1)).toEqual(["Material Detail", "SPK", "Mat, A", "RATE %", "1.00%", "-"]);
    const text = csvText([["a", "1,000"], []]);
    expect(text).toBe('﻿a,"1,000"\r\n\r\n');
  });
});
