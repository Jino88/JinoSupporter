import { describe, expect, it } from "vitest";
import {
  blankEntry, blankLog, buildTree, fieldsOf, byMonth, countByKind, entryProblem, filterTree, firstModel, isLate, isLogLate, isOpenIssue, issueFields,
  issuesBySeq, issuesOfStage, issueTree, logProblem, mediaKind, noteParts, notesInOrder, pastedImages, pathOf, replaceSome, sortEntries, suggestions,
} from "./logic";
import { PHASES, phaseName } from "./types";
import type { Entry, EntryLog, ModelGroup } from "./types";

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
  ...blankEntry(kind, date), id, seq: 0, logs: [], images: [], modelName: "M", title: id, revision: 1,
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
    expect(entryProblem({ ...f, title: "x", date: "11/09" })).toBe("Date must be a date (yyMMdd).");
    expect(entryProblem({ ...f, title: "x".repeat(201) })).toMatch(/200/);
  });
});

describe("issue register", () => {
  it("closes an issue by its done date and keeps Status in step", () => {
    const open = entry("a", "issue", "2026-11-01");
    expect(isOpenIssue(open)).toBe(true);
    expect(isOpenIssue({ ...open, doneDate: "2026-11-08" })).toBe(false);
    expect(isOpenIssue({ ...open, status: "Closed" })).toBe(false); // saved before done dates existed
    expect(isLate({ ...open, due: "2026-11-05", doneDate: "2026-11-08" }, "2026-11-09")).toBe(false);
    expect(issueFields({ ...open, doneDate: "2026-11-08" }).status).toBe("Closed");
    expect(issueFields({ ...open, doneDate: "" }).status).toBe("Open");
  });
  it("starts a new issue in the model's phase and checks the done date", () => {
    const f = { ...blankEntry("issue", "2026-11-09", "DVT"), title: "Buzz" };
    expect(f.phase).toBe("DVT");
    expect(blankEntry("dev", "2026-11-09", "DVT").phase).toBe("");
    expect(entryProblem({ ...f, doneDate: "2026-11-01" })).toMatch(/before the registered date/);
    expect(entryProblem({ ...f, doneDate: "11/10" })).toBe("Done date must be a date (yyMMdd).");
    expect(entryProblem({ ...f, doneDate: "2026-11-10" })).toBe("");
  });
  it("lists issues by number and merges used values into the suggestions", () => {
    const list = [entry("b", "issue", "2026-11-02", { seq: 2 }), entry("d", "dev", "2026-11-03", { seq: 1 }), entry("a", "issue", "2026-11-05", { seq: 1 })];
    expect(issuesBySeq(list).map((e) => e.id)).toEqual(["a", "b"]);
    expect(suggestions(["Rub & buzz", " quality ", ""], ["Quality", "Design"])).toEqual(["Quality", "Design", "Rub & buzz"]);
  });
  it("puts sub-issues under their issue, in number order, one level only", () => {
    const list = [
      entry("p2", "issue", "2026-11-01", { seq: 2 }),
      entry("c2", "issue", "2026-11-03", { seq: 5, parentId: "p1" }),
      entry("p1", "issue", "2026-11-01", { seq: 1 }),
      entry("c1", "issue", "2026-11-02", { seq: 3, parentId: "p1" }),
      entry("lost", "issue", "2026-11-02", { seq: 4, parentId: "elsewhere" }),
      entry("self", "issue", "2026-11-02", { seq: 6, parentId: "self" }),
      entry("d", "dev", "2026-11-02", { seq: 1, parentId: "p1" }),
    ];
    const t = issueTree(list);
    expect(t.top.map((e) => e.id)).toEqual(["p1", "p2", "lost", "self"]);
    expect(t.subs.get("p1")?.map((e) => e.id)).toEqual(["c1", "c2"]);
    expect(t.subs.has("p2")).toBe(false);
  });

  it("lists MP and setup-row issues under Mass Production, the rest under Development, sub-issues with their issue", () => {
    const list = [
      entry("proto", "issue", "2026-11-01", { phase: "Proto" }),
      entry("none", "issue", "2026-11-01", { phase: "" }),
      entry("mp", "issue", "2026-11-01", { phase: "MP" }),
      entry("setup", "issue", "2026-11-01", { phase: "PVT", setupRowId: "row1" }),
      entry("kid", "issue", "2026-11-01", { phase: "Proto", parentId: "mp" }),
      entry("devlog", "dev", "2026-11-01", { phase: "MP" }),
    ];
    expect(issuesOfStage(list, "dev").map((e) => e.id)).toEqual(["proto", "none"]);
    expect(issuesOfStage(list, "mp").map((e) => e.id)).toEqual(["mp", "setup", "kid"]);
  });

  it("puts back what a part of the page changed without losing the other entries", () => {
    const a = entry("a", "issue", "2026-11-01"), b = entry("b", "issue", "2026-11-01"), c = entry("c", "dev", "2026-11-01");
    const b2 = { ...b, title: "edited" }, d = entry("d", "issue", "2026-11-02");
    expect(replaceSome([a, b, c], [b], [b2, d]).map((e) => [e.id, e.title])).toEqual([["a", "a"], ["c", "c"], ["b", "edited"], ["d", "d"]]);
    expect(replaceSome([a, b, c], [a, b], [b]).map((e) => e.id)).toEqual(["c", "b"]);
  });

  it("shows the old phase names as ER / DR / DV / PV and saves the new one", () => {
    expect(["Proto", "EVT", "dvt", " PVT ", "MP", "ER", "", undefined].map(phaseName)).toEqual(["ER", "DR", "DV", "PV", "MP", "ER", "", ""]);
    expect(PHASES).toEqual(["ER", "DR", "DV", "PV", "MP"]);
    expect(fieldsOf(entry("x", "issue", "2026-11-01", { phase: "EVT" })).phase).toBe("DR");
  });

  it("checks a daily note, its schedule and done date", () => {
    const f = { ...blankLog("2026-11-09"), text: "Checked 20 pcs, 2 NG" };
    expect(logProblem({ ...f, text: "  " })).toBe("Write what was done.");
    expect(logProblem({ ...f, date: "2026/11/09" })).toBe("Date must be a date (yyMMdd).");
    expect(logProblem({ ...f, text: "x".repeat(501) })).toMatch(/500 characters/);
    expect(logProblem({ ...f, text: "two\nlines" })).toMatch(/one line/);
    expect(logProblem({ ...f, detail: "d".repeat(5001) })).toMatch(/Issue detail.*5000/);
    expect(logProblem({ ...f, text: "x".repeat(500), detail: "line 1\nline 2" })).toBe("");
    expect(logProblem({ ...f, due: "11/12" })).toBe("Due must be a date (yyMMdd).");
    expect(logProblem({ ...f, doneDate: "2026-11-08" })).toMatch(/before the note date/);
    expect(logProblem({ ...f, due: "2026-11-12", doneDate: "2026-11-10" })).toBe("");
    expect(isLogLate({ ...f, due: "2026-11-08" }, "2026-11-09")).toBe(true);
    expect(isLogLate({ ...f, due: "2026-11-08", doneDate: "2026-11-09" }, "2026-11-09")).toBe(false);
    expect(isLogLate({ ...f, due: "2026-11-09" }, "2026-11-09")).toBe(false);
  });

  it("splits a note into its one-line issue and its detail", () => {
    expect(noteParts({ text: "QM stock", detail: " order qty " })).toEqual({ line: "QM stock", detail: "order qty" });
    expect(noteParts({ text: " first \r\n second\nthird ", detail: "" })).toEqual({ line: "first", detail: "second\nthird" });
    expect(noteParts({ text: "first\nsecond", detail: "kept" })).toEqual({ line: "first", detail: "second\nkept" });
    expect(noteParts({ text: "only" })).toEqual({ line: "only", detail: "" });
  });
  it("lists notes oldest first", () => {
    const n = (id: string, date: string, createdAt: string) => ({ ...blankLog(date), id, createdAt }) as EntryLog;
    expect(notesInOrder([n("c", "2026-11-09", "2"), n("a", "2026-11-08", "9"), n("b", "2026-11-09", "1")]).map((l) => l.id)).toEqual(["a", "b", "c"]);
  });
  it("takes only pictures the server stores from a paste", () => {
    const file = (name: string, type: string, size = 1000) => ({ name, type, size });
    expect(pastedImages([file("a.txt", "text/plain")], 0)).toEqual({ files: [], problem: "" });
    expect(pastedImages([file("image.png", "image/png"), file("a.txt", "text/plain")], 0).files.map((f) => f.name)).toEqual(["image.png"]);
    expect(pastedImages([file("x.svg", "image/svg+xml")], 0).problem).toMatch(/not PNG/);
    expect(pastedImages([file("big.jpg", "image/jpeg", 10 * 1024 * 1024 + 1)], 0).problem).toMatch(/10 MB/);
    expect(pastedImages([file("c.png", "image/png")], 20).problem).toMatch(/at most 20/);
  });

  it("takes videos too, told by type or by extension, up to 500 MB each and 1 GB per save", () => {
    const MB = 1024 * 1024;
    const file = (name: string, type: string, size = 1000) => ({ name, type, size });
    expect(["a.mp4|video/mp4", "b.MOV|", "c.mkv|application/octet-stream", "d.txt|", "e.png|image/png"]
      .map((t) => t.split("|")).map(([n, ty]) => mediaKind(file(n, ty)))).toEqual(["video", "video", "video", null, "image"]);
    expect(pastedImages([file("clip.mp4", "video/mp4", 300 * MB), file("a.mkv", "")], 0).files.map((f) => f.name)).toEqual(["clip.mp4", "a.mkv"]);
    expect(pastedImages([file("big.mp4", "video/mp4", 500 * MB + 1)], 0).problem).toMatch(/500 MB/);
    expect(pastedImages([file("a.wmv", "video/x-ms-wmv")], 0).problem).toMatch(/not an MP4/);
    expect(pastedImages([file("b.mp4", "video/mp4", 400 * MB)], 1, 700 * MB).problem).toMatch(/1 GB/);
    expect(pastedImages([file("c.mp4", "video/mp4", 20 * MB)], 0).problem).toBe("");
  });
});
