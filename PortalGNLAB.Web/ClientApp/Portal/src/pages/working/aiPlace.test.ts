import { describe, expect, it } from "vitest";
import { aiPictures, applyPlace, placeItems, placeRequest, readPlaceAnswer, withNewParent } from "./aiPlace";
import { readDevPlan } from "./plan";
import type { Entry, ModelInfo } from "./types";

const info: ModelInfo = {
  customer: "Apple", phase: "Proto", owner: "Kim",
  devPlan: {
    events: [{
      id: "ev1", name: "DR build", start: "2026-11-20", end: "2026-11-25", color: "#4f46e5", rows: [], bomFrom: "",
      tasks: [{ id: "t1", name: "Material in", start: "2026-11-18", end: "2026-11-19", owner: "Lee", done: false, note: "keep" }],
    }],
  },
};

const answer = readPlaceAnswer({
  info: { customer: "Apple", build: "DR1", phase: "DR", sopTarget: "2027-03-01", bogus: "x" },
  events: [
    { id: "ev1", name: "DR build", start: "", end: "", tasks: [
      { id: "t1", name: "Material in", start: "2026-11-18", end: "2026-11-19", owner: "Lee", done: true },
      { id: "", name: "Jig ready", start: "2026-11-21", end: "", owner: "Park", done: false },
    ] },
    { id: "", name: "Reliability", start: "2026-12-01", end: "2026-12-10", tasks: [{ id: "", name: "Drop test", start: "2026-12-02", end: "", owner: "", done: false }] },
  ],
  entries: [
    { kind: "issue", stage: "dev", date: "2026-11-12", title: "Buzz noise at 1 kHz", severity: "Major", due: "2026-11-30", phase: "DR", images: [1, 0, 1, -1, "2", 1.5] },
    { kind: "issue", stage: "", phase: "MP", title: "Line 2 AOI false call" },
    { kind: "dev", title: "Changed glue to 3M" },
    { kind: "review", title: "DR review", result: "Conditional" },
    { kind: "memo", title: "dropped" },
    { kind: "dev", title: "  " },
  ],
  note: "Placed it.",
  model: "gpt · medium",
});
const all = (items: { key: string }[]) => new Set(items.map((i) => i.key));

describe("AI placement answer", () => {
  it("reads the answer leniently: known info keys, entries of a known kind with a title, MP phase means an MP issue", () => {
    expect(Object.keys(answer.info)).toEqual(["customer", "build", "phase", "sopTarget"]);
    expect(answer.entries.map((e) => [e.kind, e.stage, e.title])).toEqual([
      ["issue", "dev", "Buzz noise at 1 kHz"], ["issue", "mp", "Line 2 AOI false call"], ["dev", "", "Changed glue to 3M"], ["review", "", "DR review"],
    ]);
  });

  it("lists each item under its place, leaving out fields that do not change and an existing event with nothing of its own", () => {
    const items = placeItems(answer, info);
    expect(items.map((i) => [i.key, i.place, i.title, i.detail])).toEqual([
      ["info.build", "Model Info", "Build", "DR1"],
      ["info.phase", "Model Info", "Phase", "ER → DR"],
      ["info.sopTarget", "Model Info", "SOP target", "2027-03-01"],
      ["task.0.0", "Dev Plan", "Material in", "in DR build · 2026-11-18 – 2026-11-19 · Lee · done"],
      ["task.0.1", "Dev Plan", "Jig ready", "in DR build · 2026-11-21 · Park"],
      ["event.1", "Dev Plan", "Reliability", "2026-12-01 – 2026-12-10"],
      ["task.1.0", "Dev Plan", "Drop test", "2026-12-02"],
      ["entry.0", "Dev Issue", "Buzz noise at 1 kHz", "2026-11-12 · Major · due 2026-11-30 · 2 pictures"],
      ["entry.1", "MP Issues", "Line 2 AOI false call", ""],
      ["entry.2", "Dev History", "Changed glue to 3M", ""],
      ["entry.3", "Reviews", "DR review", "Conditional"],
    ]);
    expect(items.find((i) => i.key === "task.0.0")!.change).toBe(true);
    expect(items.find((i) => i.key === "task.1.0")!.parent).toBe("event.1");
  });

  it("puts every picked item in: fields, a task changed and one added to the existing event, a new event with its task, entries", () => {
    const { info: next, entries, pictures } = applyPlace(info, answer, all(placeItems(answer, info)), "2026-11-12");
    expect(pictures).toEqual([[1, 0], [], [], []]);
    expect(next).not.toBeNull();
    expect([next!.customer, next!.build, next!.phase, next!.sopTarget, next!.owner]).toEqual(["Apple", "DR1", "DR", "2027-03-01", "Kim"]);
    const plan = readDevPlan(next!.devPlan);
    expect(plan.events.map((e) => e.name)).toEqual(["DR build", "Reliability"]);
    const [dr, rel] = plan.events;
    expect([dr.start, dr.end]).toEqual(["2026-11-20", "2026-11-25"]);
    expect(dr.tasks.map((t) => [t.id === "t1", t.name, t.done, t.owner, t.note])).toEqual([[true, "Material in", true, "Lee", "keep"], [false, "Jig ready", false, "Park", ""]]);
    expect([rel.start, rel.end, rel.tasks.map((t) => t.name)]).toEqual(["2026-12-01", "2026-12-10", ["Drop test"]]);
    expect(entries.map((e) => [e.kind, e.date, e.status, e.phase, e.severity, e.due, e.result])).toEqual([
      ["issue", "2026-11-12", "Open", "DR", "Major", "2026-11-30", ""],
      ["issue", "2026-11-12", "Open", "MP", "", "", ""],
      ["dev", "2026-11-12", "", "", "", "", ""],
      ["review", "2026-11-12", "", "", "", "", "Conditional"],
    ]);
  });

  it("leaves out what is not picked; a new event left out takes its tasks with it", () => {
    const { info: next, entries } = applyPlace(info, answer, new Set(["task.1.0", "entry.2"]), "2026-11-12");
    expect(next).toBeNull();
    expect(entries.map((e) => e.title)).toEqual(["Changed glue to 3M"]);
    const only = applyPlace(info, answer, new Set(["task.0.1"]), "2026-11-12").info!;
    expect(only.build).toBeUndefined();
    expect(readDevPlan(only.devPlan).events.map((e) => e.tasks.map((t) => t.name))).toEqual([["Material in", "Jig ready"]]);
  });

  it("tells the server the model's fields, plan ids and entries with the list each issue is in", () => {
    const entries = [
      { id: "a", kind: "issue", seq: 1, date: "2026-11-01", title: "Noise", status: "Open", phase: "DR", setupRowId: "" },
      { id: "b", kind: "issue", seq: 2, date: "2026-11-02", title: "AOI", status: "Open", phase: "MP" },
      { id: "c", kind: "dev", seq: 1, date: "2026-11-03", title: "Glue", status: "", phase: "" },
    ] as Entry[];
    const req = placeRequest("text", "2026-11-12", info, entries);
    expect(req.info).toEqual({ customer: "Apple", build: "", phase: "ER", owner: "Kim", line: "", sopTarget: "", note: "" });
    expect(req.events.map((e) => [e.id, e.tasks.map((t) => t.id)])).toEqual([["ev1", ["t1"]]]);
    expect(req.entries.map((e) => [e.id, e.stage])).toEqual([["a", "dev"], ["b", "mp"], ["c", ""]]);
  });

  it("sends each issue's parent and saves an issue AI put under an existing one as its sub-issue", () => {
    const entries = [
      { id: "a", kind: "issue", seq: 1, date: "2026-11-01", title: "AI 과제", status: "Open", phase: "" },
      { id: "s", kind: "issue", seq: 2, date: "2026-11-02", title: "Sub", status: "Open", phase: "", parentId: "a" },
    ] as Entry[];
    expect(placeRequest("t", "2026-11-12", info, entries).entries.map((e) => [e.id, e.parentId])).toEqual([["a", ""], ["s", "a"]]);
    const a = readPlaceAnswer({ entries: [
      { kind: "issue", stage: "dev", title: "GN-06", parentId: "a" },
      { kind: "issue", stage: "dev", title: "GN-07", parentId: "gone" },
      { kind: "issue", stage: "dev", title: "Top" },
      { kind: "dev", title: "History", parentId: "a" },
    ] });
    expect(a.entries.map((e) => e.parentId)).toEqual(["a", "gone", "", ""]);
    const items = placeItems(a, info, "devissue", entries);
    expect(items.map((i) => i.detail)).toEqual(["Sub-issue of No. 1 AI 과제", "Sub-issue of an issue", ""]);
    const put = applyPlace(info, a, all(placeItems(a, info, undefined, entries)), "2026-11-12");
    expect(put.entries.map((e) => [e.title, e.parentId])).toEqual([["GN-06", "a"], ["GN-07", "gone"], ["Top", undefined], ["History", undefined]]);
  });

  it("puts a sub-issue under an issue new in the same answer: listed under it, left out with it, added after it with its id", () => {
    const a = readPlaceAnswer({ entries: [
      { kind: "issue", stage: "dev", title: "BN → GN 이동", parentId: "new:1" },
      { kind: "issue", stage: "dev", title: "GN 출장 일정" },
      { kind: "issue", stage: "dev", title: "자료 정리", parentId: "new:1" },
      { kind: "issue", stage: "dev", title: "Top" },
    ] });
    const items = placeItems(a, info, "personal");
    expect(items.map((i) => [i.title, i.parent ?? "", i.detail.split(" · ")[0]])).toEqual([
      ["BN → GN 이동", "entry.1", "Sub-issue of GN 출장 일정 (new)"], ["GN 출장 일정", "", ""],
      ["자료 정리", "entry.1", "Sub-issue of GN 출장 일정 (new)"], ["Top", "", ""],
    ]);

    // the new issue first, its sub-issues after the others, each pointing at its place among the entries added
    const put = applyPlace(info, a, all(items), "2026-11-14");
    expect(put.entries.map((e) => [e.title, e.parentId])).toEqual([
      ["GN 출장 일정", undefined], ["Top", undefined], ["BN → GN 이동", "new:0"], ["자료 정리", "new:0"],
    ]);
    expect(withNewParent(put.entries[2], ["p1", "t1"]).parentId).toBe("p1");
    expect(withNewParent(put.entries[0], ["p1"])).toBe(put.entries[0]);
    expect("parentId" in withNewParent(put.entries[2], [])).toBe(false);

    // without the new issue, a sub-issue left in on its own is added as an issue
    const alone = applyPlace(info, a, new Set(["entry.2"]), "2026-11-14");
    expect(alone.entries.map((e) => [e.title, e.parentId])).toEqual([["자료 정리", undefined]]);
  });

  it("with a target, asks for it and lists only the items of its places", () => {
    expect(placeRequest("t", "2026-11-12", info, [], [], "devissue").target).toBe("devissue");
    expect("target" in placeRequest("t", "2026-11-12", info, [])).toBe(false);
    const keys = (t?: Parameters<typeof placeItems>[2]) => placeItems(answer, info, t).map((i) => i.key);
    expect(keys("plan")).toEqual(["task.0.0", "task.0.1", "event.1", "task.1.0"]);
    expect(keys("devissue")).toEqual(["entry.0"]);
    expect(keys("issue")).toEqual(["entry.1"]);
    expect(keys("dev")).toEqual(["entry.2"]);
    expect(keys("review")).toEqual(["entry.3"]);
    expect(keys("entries")).toEqual(["entry.0", "entry.1", "entry.2", "entry.3"]);
    // My Plans: every entry the server made into a plan, under My Plans; nothing of the model
    expect(placeItems(answer, info, "personal").map((i) => [i.key, i.place])).toEqual([
      ["entry.0", "My Plans"], ["entry.1", "My Plans"], ["entry.2", "My Plans"], ["entry.3", "My Plans"],
    ]);
    expect(keys()).toHaveLength(11);
    // only the items listed are put in: a Dev Issue target does not change Model Info or Dev Plan
    const put = applyPlace(info, answer, all(placeItems(answer, info, "devissue")), "2026-11-12");
    expect(put.info).toBeNull();
    expect(put.entries.map((e) => e.title)).toEqual(["Buzz noise at 1 kHz"]);
  });

  it("takes the pictures AI can read from what was pasted, refusing a type, a size or a count it cannot send", () => {
    const f = (name: string, type: string, size = 1000) => ({ name, type, size });
    expect(aiPictures([f("a.png", "image/png"), f("notes.txt", "text/plain"), f("b.jpg", "image/jpeg")], 0))
      .toEqual({ files: [f("a.png", "image/png"), f("b.jpg", "image/jpeg")], problem: "" });
    expect(aiPictures([f("c.bmp", "image/bmp")], 0).problem).toBe("c.bmp is not PNG, JPEG, GIF or WebP.");
    expect(aiPictures([f("d.png", "image/png", 6 * 1024 * 1024)], 0).problem).toBe("d.png is larger than 5 MB.");
    expect(aiPictures([f("e.png", "image/png"), f("g.png", "image/png")], 5).problem).toBe("Up to 6 pictures can be sent.");
    expect(aiPictures([f("notes.txt", "text/plain")], 6)).toEqual({ files: [], problem: "" });
  });

  it("puts quantities into the Material Plan and Making Plan of an event and takes out the tasks moved there", () => {
    const lpr: ModelInfo = { devPlan: { events: [{
      id: "lpr", name: "LPR", start: "2026-11-18", end: "2026-11-20", color: "#4f46e5", bomFrom: "",
      tasks: [
        { id: "tf", name: "F-PCB (Free) 720ea GN 입고", start: "2026-11-11", end: "2026-11-11", owner: "", done: false, note: "" },
        { id: "ts", name: "SPK Assy 500", start: "2026-11-17", end: "2026-11-17", owner: "", done: false, note: "" },
        { id: "tk", name: "Setup", start: "2026-11-16", end: "", owner: "", done: false, note: "" },
      ],
      rows: [
        { id: "r1", name: "F-PCB-G778B", group: "material", qty: { "2026-11-05": 100 } },
        { id: "r2", name: "TAPE-G778B", group: "material", qty: { "2026-11-12": 50 } },
      ],
    }] } };
    const a = readPlaceAnswer({ events: [{
      id: "lpr", name: "", start: "", end: "", tasks: [],
      rows: [
        { id: "r1", name: "F-PCB-G778B", group: "material", qty: { "2026-11-11": 720, "11/13": 5, "2026-11-14": -1, "2026-11-15": "9" } },
        { id: "r2", name: "TAPE-G778B", group: "material", qty: { "2026-11-12": 50 } },
        { id: "", name: "SPK Assy", group: "assy", qty: { "2026-11-17": 500 } },
      ],
      removeTasks: ["tf", "ts", "nope", "tf"],
    }] });
    expect(a.events[0].rows.map((r) => [r.group, r.qty])).toEqual([["material", { "2026-11-11": 720 }], ["material", { "2026-11-12": 50 }], ["making", { "2026-11-17": 500 }]]);
    expect(a.events[0].removeTasks).toEqual(["tf", "ts", "nope"]);

    const items = placeItems(a, lpr);
    expect(items.map((i) => [i.key, i.title, i.detail, i.change, !!i.remove])).toEqual([
      ["row.0.0", "Material Plan · F-PCB-G778B", "in LPR · 11/11: 720", true, false],
      ["row.0.2", "Making Plan · SPK Assy", "in LPR · 11/17: 500", false, false],
      ["remove.0.0", "F-PCB (Free) 720ea GN 입고", "in LPR · task taken out (2026-11-11)", true, true],
      ["remove.0.1", "SPK Assy 500", "in LPR · task taken out (2026-11-17)", true, true],
    ]);

    const ev = readDevPlan(applyPlace(lpr, a, all(items), "2026-11-12").info!.devPlan).events[0];
    expect(ev.rows.map((r) => [r.id === "r1" || r.id === "r2" ? r.id : "new", r.group, r.name, r.qty])).toEqual([
      ["r1", "material", "F-PCB-G778B", { "2026-11-05": 100, "2026-11-11": 720 }],
      ["r2", "material", "TAPE-G778B", { "2026-11-12": 50 }],
      ["new", "making", "SPK Assy", { "2026-11-17": 500 }],
    ]);
    expect(ev.tasks.map((t) => t.name)).toEqual(["Setup"]);

    // a task not picked to take out stays
    const kept = readDevPlan(applyPlace(lpr, a, new Set(["row.0.0", "remove.0.0"]), "2026-11-12").info!.devPlan).events[0];
    expect(kept.tasks.map((t) => t.name)).toEqual(["SPK Assy 500", "Setup"]);
    expect(kept.rows.length).toBe(2);
  });

  it("sends the quantity plan rows of each event", () => {
    const withRows: ModelInfo = { devPlan: { events: [{ id: "e", name: "LPR", rows: [{ id: "r", name: "PLATE", group: "material", qty: { "2026-11-11": 3 } }] }] } };
    expect(placeRequest("", "2026-11-12", withRows, []).events[0].rows).toEqual([{ id: "r", name: "PLATE", group: "material", qty: { "2026-11-11": 3 } }]);
  });
});
