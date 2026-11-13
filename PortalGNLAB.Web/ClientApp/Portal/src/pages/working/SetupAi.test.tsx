import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import SetupAi from "./SetupAi";

describe("Setup AI input", () => {
  it("asks for text or pictures and shows no answer yet", () => {
    const html = renderToStaticMarkup(<SetupAi modelName="MSM-X536B" today="2026-11-12" onDone={() => {}} onError={() => {}} />);
    expect(html).toContain('aria-label="Text to turn into setup rows"');
    expect(html).toContain("MSM-X536B");
    expect(html).toContain("Make rows with AI");
    expect(html).toContain("wk-ai-text");
    expect(html).not.toContain("wk-ai-answer");
    expect(html).not.toContain('class="card');
  });
});
