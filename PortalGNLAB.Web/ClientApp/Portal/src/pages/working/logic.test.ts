import { describe, expect, it } from "vitest";
import { matchingModels } from "./PlanView";
import { blankEntry, buildTree, byMonth, countByKind, entryProblem, filterTree, firstModel, isLate, pathOf, sortEntries } from "./logic";
import type { Entry, ModelGroup } from "./types";

const g = (id: string, name: string, parentId: string | null, models: string[], sort = 0, position = 0): ModelGroup =>
  ({ id, name, parentId, models, sort, position });

// Same shape as CompanyGlossary model_groups on 2026-11-09.
const rows = [
  g("msm", "MSM", "module", ["MSM-S711B", "MSM-X316B"], 0, 1),
  g("brs", "BRS", "spk", ["BRS-161011S08ZZ"], 0, 6),
  g("sism", "SI-SM", "module", ["SI-SM-A546B"], 1, 2),
  g("spk", "SPEAKER", null, [], 1, 3),
  g("unit", "UNIT", null, ["TIU-C11-08"], 2, 4),
  g("module", "MODULE", null, [], 13, 0),
];

const entry = (id: string, kind: Entry["kind"], date: string, extra: Partial<Entry> = {}): Entry => ({
  ...blankEntry(kind, date), id, modelName: "M", title: id, revision: 1,
  createdBy: "", createdAt: date + "T00:00:00Z", updatedBy: "", updatedAt: date + "T00:00:00Z", ...extra,
});

describe("buildTree", () => {
  it("nests by parent and orders by sort then position", () => {
    const tree = buildTree(rows);
    expect(tree.map((n) => n.name)).toEqual(["SPEAKER", "UNIT", "MODULE"]);
    const module = tree[2];
    expect(module.children.map((n) => n.name)).toEqual(["MSM", "SI-SM"]);
    expect(module.total).toBe(3);
  });

  it("keeps rows with a missing parent or a cycle instead of dropping them", () => {
    const tree = buildTree([g("a", "A", "gone", ["x"]), g("b", "B", "c", ["y"]), g("c", "C", "b", ["z"])]);
    const names = (n: { name: string; children: { name: string }[] }[]): string[] => n.flatMap((x) => [x.name, ...names(x.children as never)]);
    expect(names(tree).sort()).toEqual(["A", "B", "C"]);
  });
});

describe("tree helpers", () => {
  const tree = buildTree(rows);
  it("filters models by name and drops empty groups", () => {
    const f = filterTree(tree, "x316");
    expect(f.map((n) => n.name)).toEqual(["MODULE"]);
    expect(f[0].children.map((n) => [n.name, n.models])).toEqual([["MSM", ["MSM-X316B"]]]);
    expect(filterTree(tree, "  ")).toBe(tree);
  });
  it("finds the group path and the first model", () => {
    expect(pathOf(tree, "MSM-X316B")).toEqual(["MODULE", "MSM"]);
    expect(pathOf(tree, "NOPE")).toEqual([]);
    expect(firstModel(tree)).toBe("BRS-161011S08ZZ");
  });
});

describe("entries", () => {
  const list = [
    entry("old", "dev", "2026-09-01"),
    entry("open", "issue", "2026-11-01", { due: "2026-11-05" }),
    entry("closed", "issue", "2026-10-20", { status: "Closed", due: "2026-10-01" }),
    entry("rev", "review", "2026-11-01", { createdAt: "2026-11-01T09:00:00Z" }),
  ];
  it("sorts newest first and groups by month", () => {
    const sorted = sortEntries(list);
    expect(sorted.map((e) => e.id)).toEqual(["rev", "open", "closed", "old"]);
    expect(byMonth(sorted).map((m) => [m.month, m.entries.length])).toEqual([["2026-11", 2], ["2026-10", 1], ["2026-09", 1]]);
  });
  it("counts kinds and flags only open issues past due", () => {
    expect(countByKind(list)).toEqual({ dev: 1, issue: 2, review: 1 });
    expect(list.filter((e) => isLate(e, "2026-11-09")).map((e) => e.id)).toEqual(["open"]);
    expect(isLate(list[1], "2026-11-05")).toBe(false);
  });
  it("checks the form before saving", () => {
    const f = blankEntry("issue", "2026-11-09");
    expect(f.status).toBe("Open");
    expect(entryProblem(f)).toBe("Title is required.");
    expect(entryProblem({ ...f, title: "x" })).toBe("");
    expect(entryProblem({ ...f, title: "x", date: "11/09" })).toBe("Date must be yyyy-MM-dd.");
    expect(entryProblem({ ...f, title: "x".repeat(201) })).toMatch(/200/);
  });
});

describe("matchingModels", () => {
  it("matches the Production Input model by name, ignoring case and spaces", () => {
    const m = (id: string, name: string, build: string) =>
      ({ id, name, build, processCount: 0, materialCount: 0, revision: 1, updatedAt: "", updatedBy: "" });
    const list = [m("1", "MSM-X316B", "PV2"), m("2", "msm-x316b ", "MP1"), m("3", "MSM-X316", "PV2")];
    expect(matchingModels(list, "MSM-X316B").map((x) => x.id)).toEqual(["1", "2"]);
  });
});
