import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ModelDocs from "./ModelDocs";

describe("Model Info › Documents", () => {
  const html = renderToStaticMarkup(<ModelDocs onSaved={() => {}} onOpen={() => {}} onError={() => {}}
    model={{ name: "MSM-X526", revision: 1, updatedAt: "", updatedBy: "", info: {} } as never} />);
  it("shows both sections and the AI button", () => {
    // the Work Instructions card once (not inside another card), after the AI step
    expect(html.match(/>Work Instructions</g)).toHaveLength(1);
    expect(html.indexOf("wk-docs-ai")).toBeLessThan(html.indexOf(">Work Instructions<"));
    expect(html).toContain("<h3>JIG / Machine / Tester list</h3>");
    expect(html).toContain("✦ AI: make Process and JIG/MC List");
    expect(html).toContain('accept=".xlsx,.xlsm"');
    expect(html).toContain("No file chosen");
  });
});
