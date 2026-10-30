import { describe, expect, it } from "vitest";
import { lineSplitter } from "../../bmesReportApi";
import {
  entryLevelLabel, formatElapsed, logClass, pct, selectedCount, selectedInProductGroup, toggleGroup, toggleSubEntry, usDate,
  validateReport,
} from "./reportLogic";

const none = { groupIds: [], subEntries: [] };

describe("selection", () => {
  it("toggles groups and sub entries", () => {
    const a = toggleGroup(none, 3);
    expect(a.groupIds).toEqual([3]);
    expect(toggleGroup(a, 3).groupIds).toEqual([]);
    const b = toggleSubEntry(a, "3|mid|X");
    expect(selectedCount(b)).toBe(2);
    expect(toggleSubEntry(b, "3|mid|X").subEntries).toEqual([]);
  });
  it("counts selected groups per product group", () => {
    const groups = [
      { id: 1, name: "a", productGroup: "SPK", lsCount: 1, entries: [] },
      { id: 2, name: "b", productGroup: "TWS", lsCount: 1, entries: [] },
    ];
    expect(selectedInProductGroup(groups, { groupIds: [1, 2], subEntries: [] }, "SPK")).toBe(1);
  });
  it("labels picker levels", () => {
    expect(entryLevelLabel({ key: "1|mid|X", label: "X", lsCount: 1, depth: 1 })).toBe("MODEL");
    expect(entryLevelLabel({ key: "1|sub|X|A", label: "A", lsCount: 1, depth: 2 })).toBe("SUB 1");
    expect(entryLevelLabel({ key: "1|sub|X|A", label: "A", lsCount: 1, depth: 1 })).toBe("SUB 1");
    expect(entryLevelLabel({ key: "1|sub|X|A|B", label: "B", lsCount: 1, depth: 3 })).toBe("SUB 2");
  });
});

describe("validateReport", () => {
  it("checks the period, credentials and selection in order", () => {
    expect(validateReport("2026-10-02", "2026-10-01", false, none)).toBe("End date must be on or after start date.");
    expect(validateReport("2026-10-01", "2026-10-01", false, none)).toBe("BMES credentials are not configured.");
    expect(validateReport("2026-10-01", "2026-10-01", true, none)).toBe("Select at least one sub group.");
    expect(validateReport("2026-10-01", "2026-10-02", true, { groupIds: [1], subEntries: [] })).toBeNull();
  });
});

describe("progress display", () => {
  it("formats percentages, elapsed time, log classes and dates", () => {
    expect(pct(150)).toBe("100%");
    expect(pct(33.333)).toBe("33.3%");
    expect(formatElapsed(5.25)).toBe("5.3s");
    expect(formatElapsed(125)).toBe("2m 05s");
    expect(logClass("[10:00:00] [ERROR] boom")).toBe("err");
    expect(logClass("[warn] x")).toBe("warn");
    expect(logClass("Report complete.")).toBe("ok");
    expect(logClass("fetching")).toBe("");
    expect(usDate("2026-10-30")).toBe("10/30/2026");
  });
});

describe("lineSplitter", () => {
  it("returns only complete NDJSON lines and keeps the remainder", () => {
    const s = lineSplitter();
    expect(s.feed('{"a":1}\n{"b"')).toEqual(['{"a":1}']);
    expect(s.feed(':2}\n\n')).toEqual(['{"b":2}']);
    expect(s.feed('{"c":3}')).toEqual([]);
    expect(s.rest()).toEqual(['{"c":3}']);
    expect(s.rest()).toEqual([]);
  });
});
