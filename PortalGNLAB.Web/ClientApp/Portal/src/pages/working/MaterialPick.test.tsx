import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import DevPlan, { keptTask, materialRows, TaskEdit, withMaterials } from "./DevPlan";
import MaterialPick, { pickedMaterials } from "./MaterialPick";
import { materialOptions, type MaterialItem, type MaterialList } from "./materialList";
import { readDevPlan, type PlanTask } from "./plan";
import { DEFAULT_VOCAB } from "./planVocab";

const item = (id: string, code: string, name: string, mold: MaterialItem["mold"] = ""): MaterialItem => ({ id, code, name, qty: "2", unit: "EA", mold, note: "" });
const list: MaterialList = {
  "p:A": { label: "MSM-A", rows: [item("1", "C1", "Cover", "mass"), item("2", "T1", "Tape", "sample")] },
  "p:B": { label: "MSM-B", rows: [item("3", "c1 ", "Cover"), item("4", "", "Screw")] },
};
const task = (f: Partial<PlanTask>): PlanTask => ({ id: "t", name: "GN 자재 도착", start: "", end: "", owner: "", done: false, note: "", ...f });

describe("materials to pick for a material task", () => {
  it("lists every material of the model once, with the groups it is in", () => {
    expect(materialOptions(list).map((o) => [o.key, o.item.name, o.groups])).toEqual([
      ["c:c1", "Cover", ["MSM-A", "MSM-B"]], ["c:t1", "Tape", ["MSM-A"]], ["n:screw", "Screw", ["MSM-B"]],
    ]);
  });

  it("gives the picked materials in list order, and keeps one picked before that left the list", () => {
    const before = [{ code: "OLD", name: "Old part" }, { code: "GONE", name: "Unticked" }];
    expect(pickedMaterials(list, ["n:screw", "c:c1", "c:old"], before)).toEqual([
      { code: "C1", name: "Cover" }, { code: "", name: "Screw" }, { code: "OLD", name: "Old part" },
    ]);
  });

  it("shows the materials with their mold and group, the picked ones ticked, a gone one apart", () => {
    const html = renderToStaticMarkup(<MaterialPick title="GN 자재 도착" materials={list} value={[{ code: "T1", name: "Tape" }, { code: "OLD", name: "Old part" }]} onPick={() => {}} onClose={() => {}} />);
    expect(html).toContain('role="dialog" aria-modal="true" aria-label="Pick materials"');
    expect(html).toContain("Pick materials · GN 자재 도착");
    expect(html).toMatch(/<tr class="mold-mass"><td><input type="checkbox" aria-label="Pick C1"\/><\/td><td class="mono">C1<\/td><td>Cover<\/td><td class="q">2 EA<\/td><td><span class="wk-mat-kind mass">Mass production mold<\/span><\/td><td class="muted small">MSM-A, MSM-B<\/td>/);
    expect(html).toMatch(/<tr class="mold-sample"><td><input type="checkbox" aria-label="Pick T1" checked=""\/>/);
    expect(html).toMatch(/<tr class="gone"><td><input type="checkbox" aria-label="Pick OLD" checked=""\/>[^]*No longer in the material list/);
    expect(html).toContain(">Select 2</button>");
    const none = renderToStaticMarkup(<MaterialPick title="x" materials={{}} value={[]} onPick={() => {}} onClose={() => {}} />);
    expect(none).toContain("No material in Model Info › Material yet.");
  });
});

describe("a material task keeps its materials", () => {
  it("reads them from the saved plan, leaving out empty ones", () => {
    const plan = readDevPlan({ events: [{ id: "e", name: "DV1", tasks: [
      { id: "t", name: "GN 자재 도착", code: "material-in", materials: [{ code: " C1 ", name: "Cover" }, { code: "", name: "" }, "x"] },
      { id: "u", name: "Report", materials: [] },
    ] }] });
    expect(plan.events[0].tasks.map((t) => t.materials)).toEqual([[{ code: "C1", name: "Cover" }], undefined]);
  });

  it("keeps them on a material task only, and drops the field when none is picked", () => {
    const mats = [{ code: "C1", name: "Cover" }];
    expect(keptTask(task({ code: "material-in", detail: "GN", materials: mats }), DEFAULT_VOCAB).task.materials).toEqual(mats);
    expect(keptTask(task({ code: "sample-send", materials: mats }), DEFAULT_VOCAB).task.materials).toBeUndefined();
    expect("materials" in withMaterials(task({ materials: mats }), [])).toBe(false);
  });

  it("offers + Materials when editing a material task and shows the picked ones in its row", () => {
    const edit = renderToStaticMarkup(<table><tbody><TaskEdit value={task({ code: "material-in", detail: "GN" })} vocab={DEFAULT_VOCAB} materials={list} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    expect(edit).toContain(">+ Materials</button>");
    const other = renderToStaticMarkup(<table><tbody><TaskEdit value={task({ name: "Report" })} vocab={DEFAULT_VOCAB} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    expect(other).not.toContain("Materials</button>");
    const html = renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} vocab={DEFAULT_VOCAB} model={{
      name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { materials: list, devPlan: { events: [{ id: "e", name: "DV1", start: "", end: "", color: "#2563eb", rows: [], bomFrom: "", tasks: [
        task({ id: "p", name: "GN 자재 도착", code: "material-in", detail: "GN" }),
        task({ id: "c", name: "Material in · 자재", code: "material-in", detail: "자재", parentId: "p", materials: [{ code: "C1", name: "Cover" }, { code: "", name: "Screw" }] }),
      ] }] } },
    }} />);
    expect(html).toContain('<span class="wk-dp-mats"><span class="wk-dp-mat" title="Cover">C1</span><span class="wk-dp-mat" title="Screw">Screw</span><button type="button" class="wk-link"');
    expect(html).toContain(">Change materials</button>");
    expect(html).toContain(">+ Materials</button>");
  });
});

describe("the materials picked for a material task get a sub-task row each", () => {
  const mats = (...names: string[]) => names.map((n) => ({ code: n === "Screw" ? "" : n, name: n }));
  const subs = (rows: PlanTask[]) => rows.filter((r) => r.parentId === "p").map((r) => r.name);

  it("adds a sub-task row under the task for each picked material without one, named as it is shown", () => {
    const before = task({ id: "p", code: "material-in" });
    const after = withMaterials(before, mats("C1", "Screw", "T1"));
    const kept = [task({ id: "k", name: "T1", parentId: "p" })];
    const rows = materialRows([before], before, after, kept);
    expect(rows[0]).toBe(before);
    expect(subs(rows)).toEqual(["C1", "Screw"]);
    expect(rows.slice(1).every((r) => r.id && r.id !== "p" && !r.start && !r.code)).toBe(true);
    // picking again adds none twice
    expect(subs(materialRows(rows, after, after, kept))).toEqual(["C1", "Screw"]);
  });

  it("takes out a new row of a material no longer picked unless something is typed in it, and keeps a kept one", () => {
    const before = task({ id: "p", code: "material-in", materials: mats("C1", "Screw", "T1") });
    const open = [before, task({ id: "a", name: "C1", parentId: "p" }), task({ id: "b", name: "Screw", parentId: "p", start: "2026-11-21" }), task({ id: "x", name: "", parentId: "p" })];
    const kept = [task({ id: "k", name: "T1", parentId: "p" })];
    const rows = materialRows(open, before, withMaterials(before, []), [...kept, task({ id: "b", name: "Screw", parentId: "p" })]);
    expect(rows.map((r) => r.id)).toEqual(["p", "b", "x"]);
  });

  it("adds none under a sub-task (sub-tasks are one level deep)", () => {
    const before = task({ id: "c", code: "material-in", parentId: "p" });
    expect(materialRows([before], before, withMaterials(before, mats("C1")), [])).toEqual([before]);
  });
});
