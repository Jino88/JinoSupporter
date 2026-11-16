import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { blankRequest, type RequestItem, type WorkingRequest } from "./working/request";
import RequestTotalPage from "./working/RequestTotalPage";
import { RequestTotalShell } from "./WorkingRequestTotal";

const row = (model: string, name: string, order: number, drawingNo = ""): RequestItem =>
  ({ no: 0, model, name, lineSpec: 0, stock: 0, order, drawingNo, picture: "", note: "" });
const req = (id: string, items: RequestItem[]): WorkingRequest =>
  ({ ...blankRequest("2026-11-16"), id, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "", partName: "P", items });

const list = [
  req("a", [row("MSU-L15S10-20", "RING ARRAY JIG", 1, "BJ-01"), row("MSU-L15S10-20", "VP CARRIER JIG", 74)]),
  req("b", [row("MSU-L15S10-30", "RING ARRAY JIG", 2, "bj-01"), row("MSU-L15S10-30", "RING ARRAY JIG", 4, "BJ-02")]),
];

describe("Working › Request Total page", () => {
  it("is its own page under Working", () => {
    const html = renderToStaticMarkup(<RequestTotalShell lang="" waiting={0} onLang={() => {}} initial={[]} />);
    expect(html).toContain("Working · Request Total");
    expect(html).toContain("No JIG requested yet.");
  });

  it("shows each model's total and every JIG (name and drawing no) by model, the common ones first", () => {
    const html = renderToStaticMarkup(<RequestTotalPage onError={() => {}} initial={list} />);
    expect(html).toContain('<div class="rt-model-n">MSU-L15S10-20</div><div class="rt-model-t">75</div>');
    expect(html).toContain('<div class="rt-model-n">MSU-L15S10-30</div><div class="rt-model-t">6</div>');
    expect(html).toContain("Only JIGs common to models (1)");
    // the same name with the same drawing no is one JIG asked by both models; with another drawing no it is another JIG
    expect(html).toMatch(/<tr class="rt-common"><td>1<\/td><td class="rt-name">RING ARRAY JIG<span class="rq-tag none">common<\/span><\/td><td class="mono">BJ-01<\/td><td>2<\/td><td class="rt-m">1<\/td><td class="rt-m">2<\/td><td class="rt-tot">3<\/td><td>2<\/td><\/tr>/);
    expect(html).toMatch(/<td class="rt-name">VP CARRIER JIG<\/td><td class="mono"><span class="muted">—<\/span><\/td><td>1<\/td><td class="rt-m">74<\/td><td class="rt-m"><\/td><td class="rt-tot">74<\/td>/);
    expect(html).toContain('<td class="mono">BJ-02</td>');
    expect(html).toMatch(/<tfoot><tr><td colSpan="4">Total<\/td><td class="rt-m">75<\/td><td class="rt-m">6<\/td><td class="rt-tot">81<\/td>/i);
  });

  it("can show only the JIGs common to models", () => {
    const html = renderToStaticMarkup(<RequestTotalPage onError={() => {}} initial={list} commonOnly />);
    expect(html).toContain("BJ-01");
    expect(html).not.toContain("VP CARRIER JIG");
    expect(html).not.toContain("BJ-02");
    expect(html).toMatch(/<td class="rt-tot">3<\/td><td><\/td><\/tr><\/tfoot>/);
  });
});
