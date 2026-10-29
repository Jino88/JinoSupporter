import fixture from "../test/fixtures/daily-report-v9.json";
import { describe, expect, it } from "vitest";
import type { DailyReportEnvelope } from "./contract";
import {
  directionClass,
  formatCauseDelta,
  formatPercent,
  formatPpm,
  formatSignedPpm,
  isMainProcessType,
  processTypeLabel,
} from "./format";

const typedFixture = fixture as DailyReportEnvelope;

describe("C#-compatible display formatting", () => {
  it("formats PPM, signed PPM, and percent text", () => {
    expect(formatPpm(34337.162)).toBe("34,337");
    expect(formatSignedPpm(7104)).toBe("+7,104");
    expect(formatSignedPpm(-40813)).toBe("−40,813");
    expect(formatSignedPpm(0)).toBe("0");
    expect(formatPercent(11.3733)).toBe("(+11.4%)");
    expect(formatPercent(-50.7801)).toBe("(-50.8%)");
  });

  it("preserves direction and process-type semantics", () => {
    const cause = typedFixture.snapshot!.worsened[0].topCauses[0];
    expect(formatCauseDelta(cause)).toBe("+8,841");
    expect(directionClass(cause.deltaPpm)).toBe("dr-is-worse");
    expect(processTypeLabel("FUNCTION")).toBe("FUNC");
    expect(processTypeLabel("VISUAL")).toBe("VIS");
    expect(isMainProcessType("MAIN ASSY")).toBe(true);
  });
});
