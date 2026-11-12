import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import AiTab from "./AiTab";
import { readPlaceAnswer } from "./aiPlace";
import type { WorkingModel } from "./types";

const model: WorkingModel = { name: "MSM-X536B", info: { customer: "Apple" }, revision: 3, updatedAt: "", updatedBy: "" };
const answer = readPlaceAnswer({
  info: { build: "DR1" },
  events: [{ id: "", name: "Reliability", start: "2026-12-01", end: "", tasks: [{ id: "", name: "Drop test", start: "2026-12-02" }] }],
  entries: [{ kind: "issue", stage: "dev", title: "Buzz noise", severity: "Major" }, { kind: "review", title: "DR review", result: "Fail" }],
  note: "Build to Model Info, Reliability to Dev Plan.",
  model: "gpt · medium",
});
const render = (initial?: typeof answer) => renderToStaticMarkup(
  <AiTab model={model} entries={[]} today="2026-11-12" onSaved={() => {}} onAdded={() => {}} onOpen={() => {}} onDirty={() => {}} onError={() => {}} initial={initial} />,
);

describe("AI tab", () => {
  it("asks for text to place", () => {
    const html = render();
    expect(html).toContain("Place with AI");
    expect(html).toContain('aria-label="Text to place"');
    expect(html).not.toContain("wk-ai-answer");
  });

  it("shows the answer by place, every item kept, a task under its new event", () => {
    const html = render(answer);
    expect(html).toContain("Build to Model Info, Reliability to Dev Plan.");
    expect([...html.matchAll(/<h3>([^<]+)<span class="cnt">(\d+)<\/span><\/h3>/g)].map((m) => `${m[1]} ${m[2]}`))
      .toEqual(["Model Info 1", "Dev Plan 2", "Dev Issue 1", "Reviews 1"]);
    expect(html).toMatch(/class="wk-ai-item sub"><input type="checkbox" checked=""\/><span class="wk-ai-tag">New<\/span><span class="wk-ai-title">Drop test</);
    expect(html).toContain("Put in 5 items");
    expect(html).not.toContain("wk-ai-sent");
  });
});
