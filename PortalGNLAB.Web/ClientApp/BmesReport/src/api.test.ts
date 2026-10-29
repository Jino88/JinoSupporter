import fixture from "../test/fixtures/report-v1.json";
import { describe, expect, it, vi } from "vitest";
import { fetchReport, ReportLoadError, validateReportDocument } from "./api";
import { TAB_KEYS } from "./contract";

describe("report.json adapter", () => {
  it("accepts the v1 fixture and exact eight tab keys", () => {
    const report = validateReportDocument(fixture);
    expect(Object.keys(report.tabs)).toEqual([...TAB_KEYS]);
    expect(report.schemaVersion).toBe("1.0.0");
  });

  it("rejects an unsupported schema version explicitly", () => {
    expect(() => validateReportDocument({ ...fixture, schemaVersion: "2.0.0" })).toThrowError(
      expect.objectContaining<Partial<ReportLoadError>>({ kind: "unsupported-schema" }),
    );
  });

  it.each([
    [401, "unauthorized"],
    [404, "expired"],
    [410, "expired"],
  ] as const)("maps HTTP %s to %s", async (status, kind) => {
    const fetcher = vi.fn(async () => new Response(null, { status })) as unknown as typeof fetch;
    await expect(fetchReport("/report.json", undefined, fetcher)).rejects.toMatchObject({ kind, status });
  });
});
