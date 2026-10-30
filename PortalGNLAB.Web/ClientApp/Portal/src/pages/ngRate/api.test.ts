import { afterEach, describe, expect, it, vi } from "vitest";
import { ngRateApi, RunError } from "../../ngRateApi";

/** fetch answering with NDJSON delivered in the given chunks. */
function ndjson(chunks: string[]) {
  return vi.fn(async () => {
    const body = new ReadableStream<Uint8Array>({
      start(controller) {
        for (const c of chunks) controller.enqueue(new TextEncoder().encode(c));
        controller.close();
      },
    });
    return new Response(body, { status: 200, headers: { "Content-Type": "application/x-ndjson" } });
  });
}

afterEach(() => vi.unstubAllGlobals());

describe("ngRateApi streaming", () => {
  it("reports progress lines split across chunks, then resolves with the result", async () => {
    const fetchMock = ndjson([
      '{"type":"progress","message":"Fetching"}\n{"type":"pro',
      'gress","message":"Using NG data"}\n',
      '{"type":"result","data":{"dbPath":"x.db","reports":[],"hier":[],"exportId":null}}',
    ]);
    vi.stubGlobal("fetch", fetchMock);
    const messages: string[] = [];
    const result = await ngRateApi.report("by-model", "2026-10-01", "2026-10-30", [3], ["3|mid|M"], (m) => messages.push(m));

    expect(messages).toEqual(["Fetching", "Using NG data"]);
    expect(result.dbPath).toBe("x.db");
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("/api/ngrate/report");
    expect((init.headers as Record<string, string>)["X-Requested-With"]).toBe("PortalGNLAB");
    expect(JSON.parse(init.body as string)).toEqual({
      mode: "by-model", start: "2026-10-01", end: "2026-10-30", groupIds: [3], subEntries: ["3|mid|M"],
    });
  });

  it("rejects with a RunError that carries the panel flag", async () => {
    vi.stubGlobal("fetch", ndjson(['{"type":"error","message":"Data fetch or save failed - check the log.","panel":true}\n']));
    const error = await ngRateApi.byGroup("2026-10-01", "2026-10-30", ["G"], () => {}).catch((e: unknown) => e);
    expect(error).toBeInstanceOf(RunError);
    expect((error as RunError).panel).toBe(true);
    expect((error as RunError).message).toBe("Data fetch or save failed - check the log.");
  });

  it("fails when the stream ends without a result", async () => {
    vi.stubGlobal("fetch", ndjson(['{"type":"progress","message":"x"}\n']));
    await expect(ngRateApi.byGroup("2026-10-01", "2026-10-01", ["G"], () => {})).rejects.toThrow("without a result");
  });
});
