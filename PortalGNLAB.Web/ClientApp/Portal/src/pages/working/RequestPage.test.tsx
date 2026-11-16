import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { blankRequest, type WorkingRequest } from "./request";
import RequestPage from "./RequestPage";

const saved: WorkingRequest = {
  ...blankRequest("2026-11-05"),
  id: "r1", createdBy: "jino", createdAt: "2026-11-05T09:00:00", updatedBy: "jino", updatedAt: "2026-11-05T09:00:00",
  partName: "SM-A386TOP", itemApplication: "GLONICS", usingPlace: "C2", finishDate1: "2026-11-20", deptCode: "ME-2611004",
  items: [
    { no: 1, model: "SM-A386TOP", name: "MAIN 1 JIG-Y", lineSpec: 40, stock: 0, order: 40, drawingNo: "", picture: "", note: "" },
    { no: 2, model: "SM-A386TOP", name: "DIMENSION TEST JIG(BOND)", lineSpec: 1, stock: 0, order: 1, drawingNo: "", picture: "", note: "" },
    { no: 3, model: "SM-A386TOP", name: "MAIN 2 JIG-Y", lineSpec: 10, stock: 4, order: 0, drawingNo: "", picture: "", note: "" },
  ],
};

describe("Request page", () => {
  it("has the AI box for a pasted list or pictures and the empty paper form", () => {
    const html = renderToStaticMarkup(<RequestPage today="2026-11-05" onError={() => {}} initial={[]} />);
    expect(html).toContain('aria-label="List to make the request of"');
    expect(html).toContain("Fill the form with AI");
    expect(html).toContain("PROCESSING REQUEST");
    expect(html).toContain("No request yet.");
    expect(html).toContain('value="ME GROUP"');
    expect(html).toContain('value="New request"');
    expect(html).toContain(">2026.11.05<");
    // every row of the form fills the 13 columns of the paper form
    expect(html.match(/<col /g)).toHaveLength(13);
  });

  it("opens a saved request: list line, header fields, rows with the order made from line and stock, totals", () => {
    const html = renderToStaticMarkup(<RequestPage today="2026-11-16" onError={() => {}} initial={[saved]} openId="r1" />);
    expect(html).toContain("ME-2611004 · SM-A386TOP");
    expect(html).toContain('class="rq-li on"');
    expect(html).toContain('value="SM-A386TOP"');
    expect(html).toContain('value="GLONICS"');
    expect(html).toContain(">2026.11.20<");
    expect(html).toContain('value="DIMENSION TEST JIG(BOND)"');
    expect(html).toContain('placeholder="6"');
    expect(html).toMatch(/<tr class="rq-total"><td colSpan="5">Total<\/td><td>51<\/td><td>4<\/td><td>47<\/td>/i);
    expect(html).toContain("Delete");
  });

  it("has the model above the rows and marks drawing numbers the drawing DB does not have; an empty one can be picked", () => {
    const rows = {
      ...saved,
      items: [
        { ...saved.items[0], drawingNo: "BJKP18-08-I-0000-00" },
        { ...saved.items[1], drawingNo: "BJXX99-00-Z-0000-00" },
        { ...saved.items[2], drawingNo: "" },
      ],
    };
    const html = renderToStaticMarkup(<RequestPage today="2026-11-16" onError={() => {}} initial={[rows]} openId="r1" checked={[
      { drawingNo: "BJKP18-08-I-0000-00", status: "found", match: { drawingNo: "BJKP18-08-I-0000-00", name: "CD LASER GUIDE BASE", model: "BRS-161242", category: "MC" } },
      { drawingNo: "BJXX99-00-Z-0000-00", status: "missing", match: null },
    ]} />);
    expect(html).toContain('aria-label="Model of the rows"');
    expect(html).toContain("Fill empty rows");
    expect(html).toContain('class="rq-dwg ok" title="BJKP18-08-I-0000-00 · CD LASER GUIDE BASE · BRS-161242 · MC"');
    expect(html).toContain('class="rq-dwg miss" title="Not in the JIG/MC drawing DB"');
    expect(html).toContain("1 not in the drawing DB");
    expect(html).toContain("1 without a drawing no");
    // the missing one and the empty one can be looked up; the one found cannot
    expect(html).toContain('aria-label="Find the drawing no of row 2"');
    expect(html).toContain('aria-label="Find the drawing no of row 3"');
    expect(html).not.toContain('aria-label="Find the drawing no of row 1"');
  });
});
