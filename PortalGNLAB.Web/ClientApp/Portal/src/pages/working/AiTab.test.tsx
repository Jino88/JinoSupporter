import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import AiTab from "./AiTab";
import { readPlaceAnswer, type AiTarget } from "./aiPlace";
import type { WorkingModel } from "./types";

const model: WorkingModel = { name: "MSM-X536B", info: { customer: "Apple" }, revision: 3, updatedAt: "", updatedBy: "" };
const answer = readPlaceAnswer({
  info: { build: "DR1" },
  events: [{ id: "", name: "Reliability", start: "2026-12-01", end: "", tasks: [{ id: "", name: "Drop test", start: "2026-12-02" }] }],
  entries: [{ kind: "issue", stage: "dev", title: "Buzz noise", severity: "Major" }, { kind: "review", title: "DR review", result: "Fail" }],
  note: "Build to Model Info, Reliability to Dev Plan.",
  model: "gpt · medium",
});
const render = (initial?: typeof answer, target?: AiTarget) => renderToStaticMarkup(
  <AiTab model={model} entries={[]} today="2026-11-12" onSaved={() => {}} onAdded={() => {}} onOpen={() => {}} onDirty={() => {}} onError={() => {}}
    initial={initial} target={target} />,
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

  it("as AI input of a tab: no card, says what it makes, and lists only that tab's items", () => {
    const html = render(undefined, "devissue");
    expect(html).toMatch(/^<div class="wk-ai"><div class="muted small">/);
    expect(html).not.toContain("<h2>Place with AI</h2>");
    expect(html).toContain("AI makes it into Dev Issues (problem, severity, owner, due, cause)");
    expect(html).toContain("Make with AI");

    const answered = render(answer, "devissue");
    expect([...answered.matchAll(/<h3>([^<]+)<span class="cnt">(\d+)<\/span><\/h3>/g)].map((m) => `${m[1]} ${m[2]}`)).toEqual(["Dev Issue 1"]);
    expect(answered).toContain("Put in 1 item");
  });
});
