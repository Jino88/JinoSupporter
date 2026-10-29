import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import fixture from "../test/fixtures/daily-report-v9.json";
import type { DailyReportEnvelope } from "./contract";
import { DailyReportView } from "./App";

describe("daily report static rendering", () => {
  it("renders one table with two bands and all five inline causes", () => {
    const html = renderToStaticMarkup(<DailyReportView envelope={fixture as DailyReportEnvelope} onRefresh={() => undefined} />);
    expect(html.match(/<table/g)).toHaveLength(1);
    expect(html.indexOf("WORSENED")).toBeLessThan(html.indexOf("IMPROVED"));
    expect(html.match(/class="dr-cause-row"/g)).toHaveLength(10);
    expect(html).not.toContain("DEVIATION");
    expect(html).not.toContain("full bar =");
    expect(html).toContain("zero-based Y scale");
    expect(html).toContain("CURRENT<strong>W34</strong>");
    expect(html).toContain("BASELINE<strong>W33</strong>");
    expect(html).toContain("CHANGE<strong>W34−W33</strong>");
    expect(html).not.toContain("Find model");
    expect(html).not.toContain("Grid density");
    expect(html).not.toContain("NORMAL");
    expect(html).not.toContain("COMPACT");
    expect(html).toContain("latest on the left, older periods on the right");
    expect(html).toContain("LATEST ▶ OLDER");
    expect(html).toContain("69,566 → 62,462");
    expect(html).not.toContain("LATEST ◀ OLDER");
    expect(html).toContain("FUNCTION");
    expect(html).not.toContain("title=");
  });

  it("prints the actual older baseline period and skipped-period marker", () => {
    const envelope = structuredClone(fixture) as DailyReportEnvelope;
    envelope.snapshot!.improved[0].baselineHeader = "W31";
    envelope.snapshot!.improved[0].baselineStepsBack = 3;
    const html = renderToStaticMarkup(<DailyReportView envelope={envelope} onRefresh={() => undefined} />);
    expect(html).toContain("W31");
    expect(html).toContain("−3w");
    expect(html).toContain("BASELINE<strong>W31</strong>");
    expect(html).toContain("CHANGE<strong>W34−W31</strong>");
  });
});
