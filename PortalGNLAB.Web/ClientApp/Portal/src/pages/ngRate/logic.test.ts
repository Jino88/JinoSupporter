import { describe, expect, it } from "vitest";
import type { HierSubRow, Period, ReasonRow } from "../../ngRateApi";
import {
  chartAxis,
  displayReasonRows,
  formatPpm,
  hasAnyValue,
  isHierRowVisible,
  logClass,
  maxCellPositions,
  meanOfPositive,
  parseCellNumber,
  plural,
  ppmClass,
  processTypeClass,
  processTypeColor,
  stampLog,
  summaryChartAxis,
  toggled,
  today,
  usDate,
  visibleCols,
} from "./logic";

const cols = (prefix: string, n: number): Period[] =>
  Array.from({ length: n }, (_, i) => ({ key: `${prefix}${i}`, header: `${prefix.toUpperCase()}${i}` }));

const reason = (reasonName: string, no: number, isTotal: boolean, ppm: Record<string, number> = { d0: 1 }): ReasonRow => ({
  reason: reasonName,
  no,
  isTotal,
  processType: "SUB",
  processName: "P",
  ngName: "N",
  ppm,
  groups: [],
});

const sub = (rowId: string, parentId: string | null, hasChildren = false): HierSubRow => ({
  rowId,
  parentId,
  display: rowId,
  depth: 1,
  hasChildren,
  isLineShift: false,
  ppm: {},
});

describe("formatting", () => {
  it("formats PPM as N0 with '-' for no value", () => {
    expect(formatPpm(0)).toBe("-");
    expect(formatPpm(-3)).toBe("-");
    expect(formatPpm(1234.5)).toBe("1,235");
    expect(formatPpm(999.4)).toBe("999");
    expect(ppmClass(0)).toBe("ppm-zero");
    expect(ppmClass(5)).toBe("ppm-val");
  });

  it("maps process types to badge classes like NgRateUi", () => {
    expect(processTypeClass("SUB ASSY")).toBe("pt-sub");
    expect(processTypeClass(" main")).toBe("pt-main");
    expect(processTypeClass("FUNCTION")).toBe("pt-func");
    expect(processTypeClass("VISUAL")).toBe("pt-vis");
    expect(processTypeClass("FINAL VISUAL")).toBe("pt-vis");
    expect(processTypeClass("OQC")).toBe("pt-oqc");
    expect(processTypeClass(null)).toBe("pt-other");
  });

  it("colours chart bars by process type with a cycling fallback", () => {
    expect(processTypeColor("SUB2", 0)).toBe("rgba(239,68,68,0.85)");
    expect(processTypeColor("MAIN", 3)).toBe("rgba(34,197,94,0.85)");
    expect(processTypeColor("OQC", 0)).toBe("rgba(249,115,22,0.85)");
    expect(processTypeColor("OQC", 5)).toBe("rgba(234,179,8,0.85)");
  });

  it("colours log lines", () => {
    expect(logClass("[ERROR] boom")).toBe("text-danger");
    expect(logClass("[warn] slow")).toBe("text-warning");
    expect(logClass("Fetch done")).toBe("text-success");
    expect(logClass("Report complete.")).toBe("text-success");
    expect(logClass("Fetching data")).toBe("");
  });

  it("stamps log lines and formats days", () => {
    const at = new Date(2026, 9, 30, 8, 5, 9);
    expect(stampLog("hello", at)).toBe("[08:05:09] hello");
    expect(today(at)).toBe("2026-10-30");
    expect(usDate("2026-10-30")).toBe("10/30/2026");
    expect(usDate("bad")).toBe("bad");
    expect(plural(1, "group")).toBe("1 group");
    expect(plural(3, "group")).toBe("3 groups");
  });
});

describe("columns and sets", () => {
  it("keeps the first `limit` columns until expanded", () => {
    const all = cols("d", 16);
    expect(visibleCols(all, 14, false)).toHaveLength(14);
    expect(visibleCols(all, 14, true)).toHaveLength(16);
    expect(visibleCols(cols("w", 3), 4, false)).toHaveLength(3);
  });

  it("toggles set members without mutating the input", () => {
    const start = new Set(["a"]);
    const next = toggled(start, "b");
    expect([...next]).toEqual(["a", "b"]);
    expect([...toggled(next, "a")]).toEqual(["b"]);
    expect([...start]).toEqual(["a"]);
  });

  it("detects any positive value among keys", () => {
    expect(hasAnyValue({ a: 0, b: 3 }, ["a", "b"])).toBe(true);
    expect(hasAnyValue({ a: 0, c: 3 }, ["a", "b"])).toBe(false);
  });

  it("averages positive values only", () => {
    expect(meanOfPositive([0, 10, 20])).toBe(15);
    expect(meanOfPositive([0, 0])).toBe(0);
  });
});

describe("chart axes", () => {
  const ppm = { d0: 5, d1: 0, w0: 7, m0: 9 };

  it("By Group axis puts separators only between non-empty blocks", () => {
    const axis = chartAxis({ date: cols("d", 2), week: cols("w", 1), month: cols("m", 1) });
    expect(axis.labels).toEqual(["D0", "D1", "", "W0", "", "M0"]);
    expect(axis.values(ppm)).toEqual([5, null, null, 7, null, 9]);

    const noDates = chartAxis({ date: [], week: cols("w", 1), month: [] });
    expect(noDates.labels).toEqual(["W0"]);
  });

  it("Summary axis adds a separator before weeks / months even without dates", () => {
    const axis = summaryChartAxis({ date: [], week: cols("w", 1), month: cols("m", 1) });
    expect(axis.labels).toEqual(["", "W0", "", "M0"]);
    expect(axis.values(ppm)).toEqual([null, 7, null, 9]);
  });
});

describe("reason rows", () => {
  it("keeps totals, drops empty combos and caps the number of reasons", () => {
    const rows = [
      reason("A", 0, true),
      reason("A", 1, false),
      reason("A", 2, false, { d0: 0 }),
      reason("B", 0, true),
      reason("B", 1, false),
      reason("C", 0, true),
    ];
    expect(displayReasonRows(rows).map((r) => `${r.reason}${r.no}`)).toEqual(["A0", "A1", "B0", "B1", "C0"]);
    expect(displayReasonRows(rows, 2).map((r) => `${r.reason}${r.no}`)).toEqual(["A0", "A1", "B0", "B1"]);
  });
});

describe("hierarchy rows", () => {
  it("shows a row only while every sub ancestor is expanded", () => {
    const s1 = sub("S1::a", "G::M", true);
    const s2 = sub("S2::a::b", "S1::a", true);
    const ls = sub("LS::S2::a::b::L1", "S2::a::b");
    const byId = new Map([s1, s2, ls].map((r) => [r.rowId, r]));
    expect(isHierRowVisible(s1, byId, new Set())).toBe(true);
    expect(isHierRowVisible(s2, byId, new Set())).toBe(false);
    expect(isHierRowVisible(ls, byId, new Set(["S1::a"]))).toBe(false);
    expect(isHierRowVisible(ls, byId, new Set(["S1::a", "S2::a::b"]))).toBe(true);
  });
});

describe("column maximum highlight", () => {
  it("parses cell numbers like ngRateTableStyle", () => {
    expect(parseCellNumber("1,234")).toBe(1234);
    expect(parseCellNumber("12 (3)")).toBe(12);
    expect(parseCellNumber("-")).toBeNull();
    expect(parseCellNumber(null)).toBeNull();
  });

  it("marks the column maximum, ties within 0.5 included", () => {
    const positions = maxCellPositions([
      [null, 10.2, 3],
      [null, 10, 0],
      [null, 4, 2],
    ]);
    expect(positions).toEqual([[0, 1], [1, 1], [0, 2]]);
  });
});
