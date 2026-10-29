import fixture from "../test/fixtures/report-v1.json";
import { describe, expect, it } from "vitest";
import { validateReportDocument } from "./api";
import {
  filterDailyReasonRows,
  parseStoredPreferences,
  preferencesFromDefaults,
  resolveFCostDataset,
  selectVisiblePeriods,
  sumVisiblePpmByPeriod,
} from "./logic";

const report = validateReportDocument(fixture);

describe("viewer-only pure logic", () => {
  it("uses Minimum PPM 500 and retains canonical total rows", () => {
    const daily = report.tabs.daily.data!;
    const filtered = filterDailyReasonRows(daily.modelSections[0].reasonRows, daily.referenceDatePeriodKey, 500);
    expect(filtered.map((row) => row.rowId)).toEqual(["reason-total", "reason-1"]);
    expect(preferencesFromDefaults(report.viewerDefaults).minimumPpm).toBe(500);
  });

  it("recomputes the visible detail total without mutating canonical rows", () => {
    const daily = report.tabs.daily.data!;
    const filtered = filterDailyReasonRows(daily.modelSections[0].reasonRows, daily.referenceDatePeriodKey, 500);
    const totals = sumVisiblePpmByPeriod(filtered, daily.periods);
    expect(totals["2026-08-18"]).toBe(1200);
    expect(daily.modelSections[0].reasonRows).toHaveLength(3);
  });

  it("applies date/week/month caps independently", () => {
    const daily = report.tabs.daily.data!;
    const periods = selectVisiblePeriods(daily.periods, { dateColumnLimit: 0, weekColumnLimit: 1, monthColumnLimit: 1 });
    expect(periods.map((period) => period.key)).toEqual(["W:202633", "M:202608"]);
  });

  it("resolves all follower tabs to the leader F-COST dataset", () => {
    for (const key of ["fcost-all", "fcost-weekly", "fcost-weekly-all"] as const) {
      expect(resolveFCostDataset(report, report.tabs[key].data!)).toBe(report.tabs.fcost.data!.dataset);
    }
  });

  it("falls back safely when stored viewer state is malformed", () => {
    expect(parseStoredPreferences("{broken", report.viewerDefaults)).toEqual(preferencesFromDefaults(report.viewerDefaults));
    expect(parseStoredPreferences('{"minimumPpm":750,"dateColumnLimit":2.8}', report.viewerDefaults)).toMatchObject({
      minimumPpm: 750,
      dateColumnLimit: 2,
      weekColumnLimit: 4,
      monthColumnLimit: 3,
    });
  });
});
