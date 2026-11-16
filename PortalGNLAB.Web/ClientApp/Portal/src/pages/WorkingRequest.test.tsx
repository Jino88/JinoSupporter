import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { RequestShell } from "./WorkingRequest";

describe("Working › Request page", () => {
  it("is its own page under Working with the AI box and the form", () => {
    const html = renderToStaticMarkup(<RequestShell lang="" waiting={0} onLang={() => {}} initial={[]} />);
    expect(html).toContain("Working · Request");
    expect(html).toContain("Fill the form with AI");
    expect(html).toContain("PROCESSING REQUEST");
  });
});
