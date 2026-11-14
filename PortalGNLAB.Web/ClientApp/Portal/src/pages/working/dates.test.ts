import { describe, expect, it } from "vitest";
import { WEEKDAYS, addDays, localToday, weekday } from "./dates";

describe("dates", () => {
  it("adds days across a month and a year end", () => {
    expect(addDays("2026-01-31", 1)).toBe("2026-02-01");
    expect(addDays("2027-01-01", -1)).toBe("2026-12-31");
  });

  it("names the weekday of a day", () => {
    expect(WEEKDAYS[weekday("2026-11-14")]).toBe("SAT");
  });

  it("writes today as yyyy-MM-dd", () => {
    expect(localToday()).toMatch(/^\d{4}-\d{2}-\d{2}$/);
  });
});
