import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ModelDocs from "./ModelDocs";

describe("Model Info › Documents", () => {
  const html = renderToStaticMarkup(<ModelDocs onSaved={() => {}} onOpen={() => {}} onError={() => {}}
    model={{ name: "MSM-X526", revision: 1, updatedAt: "", updatedBy: "", info: {} } as never} />);
  it("shows both sections and the AI button", () => {
    expect(html).toContain("<h3>Work Instructions</h3>");
    expect(html).toContain("<h3>JIG / Machine / Tester list</h3>");
    expect(html).toContain("✦ AI: make Process and JIG/MC List");
    expect(html).toContain('accept=".xlsx,.xlsm"');
    expect(html).toContain("No file chosen");
  });
});
