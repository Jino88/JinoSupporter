import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { moveRow, PlanVocabEditor, splitList, toRows, toVocab } from "./PlanVocabCard";
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

  it("gives every row a grip to drag it to another place", () => {
    const html = renderToStaticMarkup(<PlanVocabEditor initial={DEFAULT_VOCAB} />);
    expect(html).toContain('aria-label="Move task material-in"');
    expect(html).toContain('aria-label="Move stage');
    expect((html.match(/class="set-grip"/g) ?? []).length).toBe(DEFAULT_VOCAB.stages.length + DEFAULT_VOCAB.tasks.length);
  });

  it("moves a dragged row after the row dropped on going down, before it going up, and saves in that order", () => {
    const rows = ["a", "b", "c", "d"];
    expect(moveRow(rows, 0, 2)).toEqual(["b", "c", "a", "d"]);
    expect(moveRow(rows, 3, 1)).toEqual(["a", "d", "b", "c"]);
    expect(moveRow(rows, 1, 1)).toBe(rows);
    expect(moveRow(rows, 1, 9)).toBe(rows);
    const tasks = toRows(DEFAULT_VOCAB).tasks;
    const moved = moveRow(tasks, tasks.length - 1, 0);
    expect(toVocab([], moved).tasks.map((t) => t.code)).toEqual([tasks[tasks.length - 1].code, ...tasks.slice(0, -1).map((t) => t.code)]);
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
