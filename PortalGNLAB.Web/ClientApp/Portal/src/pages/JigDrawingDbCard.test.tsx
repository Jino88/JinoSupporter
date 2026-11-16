import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { JigDrawingDbCard } from "./Setting";

describe("Setting › JIG/MC Drawing DB", () => {
  it("shows the folder typed, the DB file it comes to, and Save / Test", () => {
    const html = renderToStaticMarkup(<JigDrawingDbCard initial={{ path: "E:\db\JigMcDrawingDB", defaultPath: "E:\db\JigMcDrawingDB", resolved: "E:\db\JigMcDrawingDB\ORIGINAL_DRAWING.db" }} />);
    expect(html).toContain("JIG/MC Drawing DB");
    expect(html).toContain('value="E:\db\JigMcDrawingDB"');
    expect(html).toContain("ORIGINAL_DRAWING.db");
    expect(html).toContain(">Test<");
  });

  it("says when no DB file is there", () => {
    const html = renderToStaticMarkup(<JigDrawingDbCard initial={{ path: "X:\none", defaultPath: "E:\db\JigMcDrawingDB", resolved: "" }} />);
    expect(html).toContain("not found");
  });
});
