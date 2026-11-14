import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import ProcessTab from "./ProcessTab";
import { blankProcess } from "./manual";
import { addSub, emptyTempGroup } from "./tempGroup";

const manualGroup = { materials: ["UPPER"], processes: [{ ...blankProcess(), name: "Top bonding", ngItems: ["Crack"], materials: [{ material: "UPPER", qtyPerUnit: 1 }] }] };
const render = (info: Record<string, unknown>) => renderToStaticMarkup(
  <ProcessTab onDirty={() => {}} onSaved={() => {}} onError={() => {}} model={{ name: "MSM-X526", revision: 1, updatedAt: "", updatedBy: "", info }} />);

describe("Process tab", () => {
  it("shows the process table of the model with its NG items and materials and a Save button", () => {
    const html = render({ manualGroup });
    expect(html).toContain("<h2>Process</h2>");
    expect(html).toContain('value="Top bonding"');
    expect(html).toContain("Crack");
    expect(html).toContain(">UPPER<");
    expect(html).toContain("Save");
    expect(html).not.toContain("Unsaved");
  });

  it("shows the table of a model linked to BMES too", () => {
    expect(render({ bmesMaterial: "MSM-X526B", manualGroup })).toContain('value="Top bonding"');
  });

  it("offers a chip per sub group of the temporary model group, read only from Model Info", () => {
    const temp = addSub(emptyTempGroup(), null, "MSM-X526-TOP");
    const html = render({ tempGroup: temp });
    expect(html).toContain("Processes and materials of");
    expect(html).toContain("MSM-X526-TOP");
  });
});
