import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { TrContext } from "./i18n";
import ManualGroup from "./ManualGroup";
import { blankProcess } from "./manual";

describe("model group table changed in place", () => {
  it("shows every process as boxes to type in, without an Edit button", () => {
    const g = {
      materials: ["UPPER", "LOWER"],
      processes: [{ ...blankProcess(), name: "AOI 1", localName: "Bôi keo", ngItems: ["Glue overflow", "Misalign"], materials: [{ material: "UPPER", qtyPerUnit: 2 }] }],
    };
    const html = renderToStaticMarkup(<ManualGroup modelName="X316B" value={g} onChange={() => {}} onError={() => {}} />);
    expect(html).toContain('value="AOI 1"');
    expect(html).toContain('value="Bôi keo"');
    expect(html).toMatch(/<textarea[^>]*wk-man-ng[^>]*>Glue overflow\nMisalign<\/textarea>/);
    expect(html).toMatch(/aria-label="Qty per unit"[^>]*value="2"|value="2"[^>]*aria-label="Qty per unit"/);
    expect(html).not.toMatch(/>Edit</);
  });

  it("shows the table in the language picked, keeping the names and NG items as written in their boxes", () => {
    const ko: Record<string, string> = { "Process name": "공정명", "NG items": "불량 항목", "Xếp UPPER": "UPPER 적재", "Glue overflow": "본드 넘침", "Ask AI": "AI에게 묻기" };
    const g = { materials: ["UPPER"], processes: [{ ...blankProcess(), name: "Xếp UPPER", ngItems: ["Glue overflow", "Misalign"], materials: [] }] };
    const html = renderToStaticMarkup(
      <TrContext.Provider value={(t) => ko[t] ?? t}>
        <ManualGroup modelName="X316B" value={g} onChange={() => {}} onError={() => {}} />
      </TrContext.Provider>);
    expect(html).toContain(">공정명</th>");
    expect(html).toContain(">불량 항목</th>");
    expect(html).toContain(">AI에게 묻기</button>");
    // what is saved stays in the boxes; the translation is under them (an item not translated yet as written)
    expect(html).toContain('value="Xếp UPPER"');
    expect(html).toMatch(/<textarea[^>]*>Glue overflow\nMisalign<\/textarea>/);
    expect(html).toContain('<div class="wk-man-tr">UPPER 적재</div>');
    expect(html).toContain('<div class="wk-man-tr"><div>본드 넘침</div><div>Misalign</div></div>');
    // nothing extra when the language is the original
    expect(renderToStaticMarkup(<ManualGroup modelName="X316B" value={g} onChange={() => {}} onError={() => {}} />)).not.toContain("wk-man-tr");
  });
});
