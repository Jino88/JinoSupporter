import fixture from "../test/fixtures/daily-report-v9.json";
import { describe, expect, it } from "vitest";
import { DailyReportContractError, validateDailyReportEnvelope } from "./contract";

describe("daily report contract", () => {
  it("accepts the schema 9 fixture", () => {
    const envelope = validateDailyReportEnvelope(fixture);
    expect(envelope.snapshot?.schemaVersion).toBe(9);
    expect(envelope.snapshot?.worsened[0].modelName).toBe("BRS-161011S08ZZ_B2");
  });

  it("rejects a mismatched schema so the Razor host can fall back", () => {
    const mismatched = structuredClone(fixture);
    if (mismatched.snapshot) mismatched.snapshot.schemaVersion = 8;
    expect(() => validateDailyReportEnvelope(mismatched)).toThrowError(DailyReportContractError);
    try {
      validateDailyReportEnvelope(mismatched);
    } catch (error) {
      expect(error).toMatchObject({ kind: "unsupported-schema" });
    }
  });
});
