import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, vi } from "vitest";
import { AnalysisCard } from "./WorkingGraphMaker";
import { analyzeAndRender } from "./workingGraph/analyze";
import { parseTable } from "./graphMaker/data";

const text = "Time\tTemp\tPress\n1\t20\t5\n2\t21\t6\n3\t22\t7";
const table = parseTable(text, "auto", 1, false);
const input = { text, delimiter: "auto", headerRowNumber: 1, headerNotUse: false };
const json = (body: unknown, status = 200) =>
  new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json" } });

describe("Working › Graphmaker", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("renders with the kind and options the AI picked", async () => {
    const calls: { url: string; body: any }[] = [];
    vi.stubGlobal("fetch", vi.fn(async (url: string, init: RequestInit) => {
      calls.push({ url, body: JSON.parse(String(init.body)) });
      if (url.endsWith("working-graph/analyze"))
        return json({ kind: 1, title: "Temp vs Press", reason: "Two numeric columns", options: { xColumns: ["Temp"], yColumns: ["Press"] }, model: "gpt-x", elapsedSeconds: 42 });
      return json({ kind: "Scatter" });
    }));

    const seen: number[] = [];
    const { result, options, payloads } = await analyzeAndRender(input, table, undefined, (r) => seen.push(r.kind));

    expect(calls.map((c) => c.url)).toEqual(["/api/portal/tools/working-graph/analyze", "/api/portal/tools/working-graph/render"]);
    expect(calls[0].body).toEqual(input);
    expect(calls[1].body).toMatchObject({ text, delimiter: "auto", headerRowNumber: 1, headerNotUse: false });
    expect(calls[1].body.request).toMatchObject({ graphType: 1, xColumns: ["Temp"], yColumns: ["Press"] });
    expect(options.xColumns).toEqual(["Temp"]);
    expect(result.kind).toBe(1);
    expect(seen).toEqual([1]);
    expect(payloads).toHaveLength(1);
  });

  it("surfaces an analyze error and never renders", async () => {
    const fetchMock = vi.fn(async () => json({ error: "CLI failed" }, 400));
    vi.stubGlobal("fetch", fetchMock);
    await expect(analyzeAndRender(input, table)).rejects.toThrow("CLI failed");
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });

  it("shows the AI's pick, reason, model and time", () => {
    const html = renderToStaticMarkup(<AnalysisCard analysis={{ kind: 3, title: "Cpk of Temp", reason: "Spec limits present", options: {}, model: "gpt-x", elapsedSeconds: 42 }} />);
    expect(html).toContain("AI pick: CPK");
    expect(html).toContain("Cpk of Temp");
    expect(html).toContain("Spec limits present");
    expect(html).toContain("gpt-x");
  });
});
