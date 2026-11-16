import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { PlanVocabEditor, splitList, toRows, toVocab } from "./PlanVocabCard";
import { DEFAULT_VOCAB } from "./working/planVocab";

describe("Setting › Dev Plan standard names", () => {
  it("shows the stage and task tables with Save and Reset to default", () => {
    const html = renderToStaticMarkup(<PlanVocabEditor initial={DEFAULT_VOCAB} />);
    expect(html).toContain("Stages");
    expect(html).toContain("Tasks");
    expect(html).toContain('value="material-in"');
    expect(html).toContain("자재 도착, 자재도착");
    expect(html).toContain("design-confirm, mold-order");
    expect(html).toContain(">Add stage<");
    expect(html).toContain(">Add task<");
    expect(html).toContain(">Save<");
    expect(html).toContain(">Reset to default<");
  });

  it("splits comma-separated text and drops blanks", () => {
    expect(splitList(" a, b ,, c;d\ne ")).toEqual(["a", "b", "c", "d", "e"]);
    expect(splitList("")).toEqual([]);
  });

  it("round-trips the default list through the table rows", () => {
    const rows = toRows(DEFAULT_VOCAB);
    expect(toVocab(rows.stages, rows.tasks)).toEqual(DEFAULT_VOCAB);
  });

  it("builds the template of each stage from its text and skips stages without a code", () => {
    const v = toVocab(
      [
        { code: " ER ", name: "ER", aliases: "", template: "a, b" },
        { code: "", name: "x", aliases: "", template: "a" },
      ],
      [{ code: "a", name: "A", aliases: "x, y", template: "" }],
    );
    expect(v.templates).toEqual({ ER: ["a", "b"] });
    expect(v.stages[0].code).toBe("ER");
    expect(v.tasks[0].aliases).toEqual(["x", "y"]);
  });
});
