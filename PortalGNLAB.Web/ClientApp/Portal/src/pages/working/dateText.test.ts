import { describe, expect, it } from "vitest";
import { dayText, readDayText } from "./dateText";

describe("dates typed as yyMMdd", () => {
  it("reads yyMMdd, yyyyMMdd and dates with separators as yyyy-MM-dd", () => {
    expect(readDayText("261112")).toBe("2026-11-12");
    expect(readDayText(" 270105 ")).toBe("2027-01-05");
    expect(readDayText("20261112")).toBe("2026-11-12");
    expect(readDayText("2026-11-12")).toBe("2026-11-12");
    expect(readDayText("26/1/5")).toBe("2026-01-05");
    expect(readDayText("2028.02.29")).toBe("2028-02-29");
    expect(readDayText("")).toBe("");
  });

  it("refuses what is not a day", () => {
    for (const bad of ["2611", "261131", "270229", "261312", "11/12", "abc", "2611120"]) expect(readDayText(bad)).toBeNull();
  });

  it("shows a date as yyMMdd and leaves anything else as typed", () => {
    expect(dayText("2026-11-12")).toBe("261112");
    expect(dayText("")).toBe("");
    expect(dayText("2611")).toBe("2611");
  });
});
