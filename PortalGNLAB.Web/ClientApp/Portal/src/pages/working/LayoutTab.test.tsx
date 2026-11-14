import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import LayoutTab from "./LayoutTab";
import type { Layout } from "./layout";

const layout: Layout = {
  fileName: "line.dxf", size: 3145728, uploadedAt: "2026-11-14T09:30:00Z", uploadedBy: "kim",
  bounds: [0, 0, 10, 10], layers: ["WALL", "TEXT"],
  paths: [{ l: 0, c: null, p: [0, 0, 10, 10] }], texts: [{ l: 1, c: "#ff0000", x: 1, y: 1, h: 2, r: 0, t: "A" }],
};
const render = (initial: Layout | null | undefined) =>
  renderToStaticMarkup(<LayoutTab modelName="MSM-X316B" initial={initial} onError={() => {}} />);

describe("process layout tab", () => {
  it("asks for a DXF when the model has no layout", () => {
    const html = render(null);
    expect(html).toContain("<h2>Process Layout</h2>");
    expect(html).toContain("Upload DXF");
    expect(html).toContain('accept=".dxf"');
    expect(html).not.toContain("<canvas");
  });

  it("shows the file, buttons, layer chips and a canvas when it has one", () => {
    const html = render(layout);
    expect(html).toContain("line.dxf");
    expect(html).toContain("3.0 MB");
    expect(html).toContain("2026-11-14 09:30");
    expect(html).toContain("kim");
    for (const label of [">Fit<", ">Replace<", ">Delete<"]) expect(html).toContain(label);
    expect(html).toMatch(/<button class="wk-chip on" title="WALL">WALL<\/button>/);
    expect(html).toMatch(/<button class="wk-chip on" title="TEXT">TEXT<\/button>/);
    expect(html).toContain('<canvas class="wk-lay-canvas"');
    expect(html).not.toContain("Upload DXF");
  });

  it("waits while the layout is being read", () => {
    expect(render(undefined)).toContain("Loading");
  });
});
