import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
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
});
