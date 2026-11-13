import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import AiPopup from "./AiPopup";

describe("AI input popup", () => {
  it("shows the title, a close button and the AI box", () => {
    const html = renderToStaticMarkup(<AiPopup title="AI input · Dev Plan" dirty={false} onClose={() => {}}><p>box</p></AiPopup>);
    expect(html).toContain('role="dialog" aria-modal="true" aria-label="AI input · Dev Plan"');
    expect(html).toContain('<h2>AI input · Dev Plan</h2><button class="btn sm" aria-label="Close">×</button>');
    expect(html).toContain('<div class="wk-ai-pop-b"><p>box</p></div>');
  });
});
