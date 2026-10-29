import fixture from "../test/fixtures/daily-report-v9.json";
import { describe, expect, it } from "vitest";
import type { DailyReportEnvelope, DailyReportModelRow } from "./contract";
import {
  TREND_DIRECTION,
  baselineGapLabel,
  buildSparkGeometry,
  filterAndSortRows,
} from "./logic";

const snapshot = (fixture as DailyReportEnvelope).snapshot!;

describe("daily report presentation logic", () => {
  it("filters and sorts without mutating the service values", () => {
    const rows = [...snapshot.worsened, ...snapshot.improved];
    const result = filterAndSortRows(rows, { group: "SPK", query: "msu", sort: "delta" });
    expect(result.map((row) => row.modelName)).toEqual(["MSU-L20S15-07_E2"]);
    expect(rows).toHaveLength(2);
  });

  it("plots newest data on the left and breaks the line at missing zero periods", () => {
    const geometry = buildSparkGeometry([500, 0, 300, 200], 300)!;
    expect(TREND_DIRECTION).toBe("latest-left");
    expect(geometry.latestPoint.sourceIndex).toBe(0);
    expect(geometry.latestPoint.x).toBe(3.5);
    expect(geometry.segments.map((segment) => segment.map((point) => point.sourceIndex))).toEqual([[0], [2, 3]]);
    expect(geometry.zeroY).toBe(26.5);
    expect(geometry.baselineY).toBeCloseTo(12.7, 5);
    expect(geometry.segments[1][1].y).toBeCloseTo(17.3, 5);
  });

  it("labels a production-aware baseline that skipped empty periods", () => {
    const row = { ...snapshot.improved[0], baselineStepsBack: 3 } as DailyReportModelRow;
    expect(baselineGapLabel(row.baselineStepsBack, "w")).toBe("−3w");
    expect(baselineGapLabel(1, "d")).toBeNull();
  });
});
