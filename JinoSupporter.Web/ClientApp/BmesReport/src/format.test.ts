import { describe, expect, it } from "vitest";
import { formatPercent, formatPpm, formatShareRatio, formatUsd, formatVnd, roundHalfToEven } from "./format";

describe("display formatting", () => {
  it("formats PPM, percent, USD, VND and preserves null vs zero", () => {
    expect(formatPpm(1234.4)).toBe("1,234");
    expect(formatPpm(0)).toBe("0");
    expect(formatPpm(null)).toBe("-");
    expect(formatPercent(1.234)).toBe("1.23%");
    expect(formatShareRatio(0.125)).toBe("12.50%");
    expect(formatUsd(12.6)).toBe("$13");
    expect(formatUsd(12.345, true)).toBe("$12.34");
    expect(formatVnd(25000)).toBe("25,000 VND");
  });

  it("matches the server midpoint-to-even display rounding", () => {
    expect(roundHalfToEven(2.5)).toBe(2);
    expect(roundHalfToEven(3.5)).toBe(4);
    expect(roundHalfToEven(-2.5)).toBe(-2);
    expect(formatPpm(2.5)).toBe("2");
  });
});
