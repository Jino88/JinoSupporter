import { describe, expect, it, vi } from "vitest";
import fixture from "../test/fixtures/daily-report-v9.json";
import { fetchDailyReport, refreshDailyReport } from "./api";

describe("daily report API", () => {
  it("uses same-origin private GET semantics", async () => {
    const fetcher = vi.fn(async () => new Response(JSON.stringify(fixture), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    })) as unknown as typeof fetch;
    const result = await fetchDailyReport("/api/bmes/daily-report", undefined, fetcher);
    expect(result.snapshot?.schemaVersion).toBe(9);
    expect(fetcher).toHaveBeenCalledWith("/api/bmes/daily-report", expect.objectContaining({
      method: "GET",
      credentials: "same-origin",
      cache: "no-store",
    }));
  });

  it("posts a refresh and maps permission failures", async () => {
    const ok = vi.fn(async () => new Response(JSON.stringify({ ...fixture, isRefreshing: true }), { status: 200 })) as unknown as typeof fetch;
    expect((await refreshDailyReport("/api/bmes/daily-report/refresh", undefined, ok)).isRefreshing).toBe(true);
    expect(ok).toHaveBeenCalledWith("/api/bmes/daily-report/refresh", expect.objectContaining({ method: "POST" }));

    const forbidden = vi.fn(async () => new Response(null, { status: 403 })) as unknown as typeof fetch;
    await expect(fetchDailyReport("/api/bmes/daily-report", undefined, forbidden)).rejects.toMatchObject({
      kind: "forbidden",
      status: 403,
    });
  });
});
