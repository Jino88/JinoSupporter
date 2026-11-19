import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { addDays } from "./dates";
import DevPlan, { afterSave, autoSaveDue, editGroup, EventEdit, Gantt, keptTask, openRows, TaskEdit, TaskEditRow, TidyNames, tint, withRows } from "./DevPlan";
import { tidyRows } from "./planNames";
import { DEFAULT_VOCAB, type PlanVocab } from "./planVocab";
import type { DevPlan as Plan, PlanEvent, PlanTask } from "./plan";

const event = (f: Partial<PlanEvent>): PlanEvent =>
  ({ id: "e", name: "DR build", start: "", end: "", color: "#2563eb", tasks: [], rows: [], bomFrom: "", ...f });

const boxes = (html: string) => [...html.matchAll(/<input class="inp wk-datebox sm" placeholder="yyMMdd"[^>]*aria-label="(Start|End)"[^>]*value="(\d*)"/g)].map((m) => `${m[1]} ${m[2]}`);

describe("event dates typed as yyMMdd", () => {
  it("shows the event Start and End as yyMMdd boxes, not the browser's date box", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({ start: "2027-01-05", end: "2028-02-28" })} onKeep={() => {}} onCancel={() => {}} />);
    expect(html).not.toContain('type="date"');
    expect(boxes(html)).toEqual(["Start 270105", "End 280228"]);
  });

  it("leaves the boxes of empty event dates empty", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({})} onKeep={() => {}} onCancel={() => {}} />);
    expect(boxes(html)).toEqual(["Start ", "End "]);
  });
});

describe("calendar with an event's quantities open", () => {
  const rows = [{ id: "a", label: "LPR", sub: "", color: "#4f46e5", span: ["2026-11-17", "2026-11-20"] as [string, string] }];
  const days = Array.from({ length: 45 }, (_, i) => addDays("2026-11-09", i));
  it("widens the day columns for quantities (scrolls) instead of fitting a short calendar to the width", () => {
    const shut = renderToStaticMarkup(<Gantt rows={rows} days={days} />);
    expect(shut).toContain('class="wk-gantt fill"');
    const open = renderToStaticMarkup(<Gantt rows={rows} days={days} selected="a" expand={<tr><td>qty</td></tr>} />);
    expect(open).toContain('class="wk-gantt qopen"');
    expect(open).not.toContain(" fill");
  });

  it("names a month of a day or two by its month only, so it does not run into the next", () => {
    const html = renderToStaticMarkup(<Gantt rows={rows} days={["2026-10-31", ...days.slice(0, 5)]} />);
    expect([...html.matchAll(/<th colSpan="(\d+)" class="mon" title="([^"]+)">([^<]+)<\/th>/g)].map((m) => `${m[2]}:${m[1]}:${m[3]}`))
      .toEqual(["2026-10:1:10", "2026-11:5:2026-11"]);
  });
});

describe("sub-tasks in the detailed schedule", () => {
  const t = (id: string, name: string, start: string, end: string, parentId?: string) =>
    ({ id, name, start, end, owner: "", done: false, note: "", ...(parentId ? { parentId } : {}) });
  const html = renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} model={{
    name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { devPlan: { events: [{ ...event({ name: "LPR" }), tasks: [
      t("a", "SETUP", "", ""), t("b", "GN parts", "2026-11-27", "2026-11-27"), t("a1", "Line A", "2026-11-30", "2026-12-01", "a"), t("a2", "Line B", "2026-12-02", "", "a"),
    ] }] } },
  }} />);
  const table = html.slice(html.indexOf("wk-dp-tasks"));

  it("lists each sub-task right under its task, marked and indented", () => {
    expect([...table.matchAll(/<tr( class="[^"]*")?[^>]*><td class="wk-dp-grip"[^]*?<td class="l">(?:<span class="wk-dp-sub">↳<\/span>)?<b>([^<]+)<\/b>/g)]
      .map((m) => `${m[2]}${m[1]?.includes("sub") ? " (sub)" : ""}`)).toEqual(["SETUP", "Line A (sub)", "Line B (sub)", "GN parts"]);
  });

  it("offers + Sub on a task only, not on a sub-task", () => {
    expect([...table.matchAll(/>\+ Sub</g)]).toHaveLength(2);
    expect(table).toMatch(/Line B<\/b>(?:(?!<tr)[^])*?<td class="act"><button class="wk-link">Edit</);
  });

  it("offers Edit all under the schedule, and Edit on a task opens it with its sub-tasks", () => {
    expect(table).toContain(">Edit all</button>");
    expect(table).toContain('title="Edit this task and its sub-tasks together">Edit</button>');
    const tasks = [t("a", "SETUP", "", ""), t("b", "GN parts", "", ""), t("a1", "Line A", "", "", "a"), t("a2", "Line B", "", "", "a")];
    expect(editGroup(tasks, "a").map((x) => x.id)).toEqual(["a", "a1", "a2"]);
    expect(editGroup(tasks, "a2").map((x) => x.id)).toEqual(["a2"]);
  });

  it("keeps Edit, + Sub and Delete on for the other rows, and + Task and Edit all under the schedule", () => {
    expect(table).not.toMatch(/class="wk-link[^"]*" disabled/);
    expect(table).toMatch(/<div class="wk-dp-adds"><button class="btn sm wk-dp-add">\+ Task<\/button>/);
  });

  it("opens rows at once: a new sub-task under its task while the task is being changed, and more on top without losing what is typed", () => {
    const tasks = [t("a", "JIG/MC Receive", "2026-11-01", "2026-11-28"), t("a1", "Line A", "", "", "a"), t("b", "Report", "", "")];
    const sub = t("n", "Rack", "", "", "a");
    const typed = { ...tasks[0], name: "JIG/MC Receive 2" };
    const open = openRows(openRows([sub], [typed]), editGroup(tasks, "a"));
    expect(open.map((x) => `${x.id} ${x.name}`)).toEqual(["n Rack", "a JIG/MC Receive 2", "a1 Line A"]);
    expect(withRows(tasks, open).map((x) => `${x.id} ${x.name}`)).toEqual(["a JIG/MC Receive 2", "a1 Line A", "n Rack", "b Report"]);
    expect(withRows(tasks, [t("z", "New", "", "")]).map((x) => x.id)).toEqual(["a", "a1", "b", "z"]);
  });

  it("gives each open row its own Cancel and Keep, Keep off while that row has a problem", () => {
    const row = (v: ReturnType<typeof t>) => renderToStaticMarkup(<table><tbody><TaskEditRow value={v} vocab={DEFAULT_VOCAB} onChange={() => {}} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    const ok = row(t("n", "Rack", "", "", "a"));
    expect(ok).toMatch(/aria-label="Task name" value="Rack"/);
    expect(ok).toMatch(/>Cancel<\/button> <button class="btn sm primary">Keep<\/button><\/td><\/tr>/);
    expect(row({ ...t("n", "Rack", "2026-12-05", "2026-12-01", "a") })).toMatch(/<button class="btn sm primary" disabled="">Keep<\/button>/);
    expect(keptTask(t("a1", "Line A ", "", "", "a"), DEFAULT_VOCAB).task.name).toBe("Line A");
  });

  it("puts the sub-tasks on the calendar under their task, the task over their days when it has none", () => {
    const cal = html.slice(html.indexOf("Detailed schedule"), html.indexOf("wk-dp-tasks"));
    expect([...cal.matchAll(/<td class="nm( child)?">[^]*?<b class="">([^<]+)<\/b><small>([^<]*)<\/small>/g)].map((m) => `${m[2]}${m[1] ?? ""} ${m[3]}`))
      .toEqual(["SETUP 11-30 – 12-02", "Line A child 11-30 – 12-01", "Line B child 12-02", "GN parts 11-27"]);
  });

  it("draws a sub-task smaller, as a thin bar in a lighter shade of the event color", () => {
    const rows = [
      { id: "a", label: "JIG/MC", sub: "", color: "#4f46e5", span: ["2026-11-17", "2026-11-18"] as [string, string] },
      { id: "a1", label: "JIG in", sub: "", color: "#4f46e5", span: ["2026-11-18", "2026-11-18"] as [string, string], child: true },
    ];
    const cal = renderToStaticMarkup(<Gantt rows={rows} days={["2026-11-17", "2026-11-18"]} />);
    const [task, sub] = cal.split("<tr").filter((x) => x.includes('<td class="nm'));
    expect(task).toContain('style="background:#4f46e5"');
    expect(task).not.toContain('class="bar"');
    expect(sub).toMatch(/^ class="[^"]*\bchild\b/);
    expect(sub).toContain(`<i class="bar" style="background:${tint("#4f46e5", 0.45)}"></i>`);
    expect(sub).not.toContain("#4f46e5");
    expect(tint("#4f46e5", 0.45)).toBe("#9e99f1");
    expect(tint("#000", 1)).toBe("#ffffff");
    expect(tint("red", 0.5)).toBe("red");
  });
});

describe("actual days in the detailed schedule", () => {
  const html = renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} model={{
    name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { devPlan: { events: [{ ...event({ name: "LPR" }), tasks: [
      { id: "a", name: "Jig ready", start: "2026-11-10", end: "2026-11-12", owner: "", done: true, note: "", actualStart: "2026-11-11", actualEnd: "2026-11-14" },
      { id: "b", name: "Report", start: "2026-11-19", end: "2026-11-20", owner: "", done: false, note: "" },
    ] }] } },
  }} />);
  const table = html.slice(html.indexOf("wk-dp-tasks"));

  it("has Actual start and Actual end beside the planned days, an end after the plan in red", () => {
    expect(table).toMatch(/<th class="fit">Start<\/th><th class="fit">End<\/th><th class="fit">Actual start<\/th><th class="fit">Actual end<\/th>/);
    expect(table).toContain('<td>2026-11-12</td><td class="wk-dp-real">2026-11-11</td><td class="wk-dp-real late">2026-11-14</td>');
    expect(table).toContain('<td>2026-11-20</td><td class="wk-dp-real">—</td><td class="wk-dp-real">—</td>');
  });
});

describe("standard names in the Dev Plan", () => {
  const sel = (html: string, label: string) => new RegExp(`<select[^>]*aria-label="${label}"[^>]*>[^]*?</select>`).exec(html)?.[0] ?? "";
  it("picks an event's stage from the list with its quantity, or a name of its own as Other", () => {
    const er = renderToStaticMarkup(<EventEdit value={event({ name: "ER (1.2k)", stage: "ER", qty: "1.2k" })} onKeep={() => {}} onCancel={() => {}} />);
    expect(sel(er, "Stage")).toMatch(/<option value="ER" selected="">ER<\/option>[^]*<option value="">Other \(own name\)<\/option>/);
    expect(er).toMatch(/aria-label="Quantity" value="1.2k"/);
    expect(er).not.toContain('aria-label="Event name"');
    const own = renderToStaticMarkup(<EventEdit value={event({ name: "DR build" })} onKeep={() => {}} onCancel={() => {}} />);
    expect(sel(own, "Stage")).toContain('<option value="" selected="">Other (own name)</option>');
    expect(own).toMatch(/aria-label="Event name" value="DR build"/);
  });

  it("picks a task from the standard tasks with a detail, or a name of its own", () => {
    const t = { id: "t", name: "Sample shipment · 300ea", start: "", end: "", owner: "", done: false, note: "", code: "sample-send", detail: "300ea" };
    const std = renderToStaticMarkup(<table><tbody><TaskEdit value={t} vocab={DEFAULT_VOCAB} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    expect(sel(std, "Category")).toContain('<option value="sample-send" selected="">Sample shipment</option>');
    expect(std).toMatch(/aria-label="Detail" value="300ea"/);
    const own = renderToStaticMarkup(<table><tbody><TaskEdit value={{ ...t, name: "Report", code: undefined, detail: undefined }} vocab={DEFAULT_VOCAB} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    expect(sel(own, "Category")).toContain('<option value="" selected="">Other (own name)</option>');
    expect(own).toMatch(/aria-label="Task name" value="Report"/);
  });

  const plan = { events: [{ ...event({ id: "e1", name: "DV1 (1.2k)", stage: "DV1", qty: "1.2k" }), tasks: [
    { id: "a", name: "Line ready", start: "", end: "", owner: "", done: false, note: "", code: "line-ready" },
    { id: "b", name: "GN 자재 도착", start: "", end: "", owner: "", done: false, note: "" },
  ] }] };
  const html = renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} vocab={DEFAULT_VOCAB}
    model={{ name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { devPlan: plan } }} />);

  it("offers the standard tasks of the event's stage it does not have yet", () => {
    expect(html).toMatch(/title="Material in, Sample making, Sample shipment">\+ Standard tasks of DV1 \(3\)<\/button>/);
  });

  it("counts the names of their own and opens their tidy-up", () => {
    expect(html).toMatch(/>Standard names <span class="cnt">1<\/span><\/button>/);
    const rows = tidyRows(plan as never, DEFAULT_VOCAB);
    const t = renderToStaticMarkup(<TidyNames rows={rows} vocab={DEFAULT_VOCAB} onRows={() => {}} onApply={() => {}} onCancel={() => {}} />);
    expect(t).toContain("<td class=\"l\">GN 자재 도착</td>");
    expect(sel(t, "Standard name of GN 자재 도착")).toContain('<option value="material-in" selected="">Material in</option>');
    expect(t).toMatch(/aria-label="Detail of GN 자재 도착" value="GN"/);
    expect(t).toContain(">Apply 1</button>");
  });
});

describe("saving the plan by itself", () => {
  const saved: Plan = { events: [event({ id: "e1", name: "LPR" })] };
  const changed: Plan = { events: [{ ...saved.events[0], tasks: [{ id: "a", name: "JIG/MC", start: "", end: "", owner: "", done: false, note: "" }] }] };

  it("saves a changed plan without Save, but not an unchanged one, one already saving or the one that just failed", () => {
    expect(autoSaveDue(changed, saved, false, "")).toBe(true);
    expect(autoSaveDue({ events: [{ ...saved.events[0] }] }, saved, false, "")).toBe(false);
    expect(autoSaveDue(changed, saved, true, "")).toBe(false);
    expect(autoSaveDue(changed, saved, false, JSON.stringify(changed))).toBe(false);
  });

  it("keeps a change made while saving, to save it next", () => {
    const kept: Plan = JSON.parse(JSON.stringify(changed));
    expect(afterSave(changed, changed, kept)).toBe(kept);
    const later: Plan = { events: [] };
    expect(afterSave(later, changed, kept)).toBe(later);
  });

  it("no longer says a change waits for Save", () => {
    const t = renderToStaticMarkup(<TidyNames rows={[]} vocab={DEFAULT_VOCAB} onRows={() => {}} onApply={() => {}} onCancel={() => {}} />);
    expect(t).not.toContain("Not saved until Save.");
  });
});

describe("task category and samples counted day by day", () => {
  const t = (id: string, name: string, start: string, end: string, f: Partial<PlanTask> = {}): PlanTask =>
    ({ id, name, start, end, owner: "", done: false, note: "", ...f });
  const tasks: PlanTask[] = [
    t("a", "Setup", "2026-12-14", "2026-12-14", { code: "setup" }),
    t("b", "Sample making · 400pcs", "2026-12-15", "2026-12-17", { code: "sample-make", detail: "400pcs", qty: { "2026-12-15": 100, "2026-12-16": 300 } }),
    t("c", "GN→AS 샘플발송", "2026-12-19", "2026-12-19", { code: "sample-send", qty: { "2026-12-16": 50, "2026-12-19": 300 } }),
    t("d", "개발라인 준비", "2026-12-13", "2026-12-13"),
  ];
  const render = (list: unknown[]) => renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} vocab={DEFAULT_VOCAB}
    model={{ name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { devPlan: { events: [{ ...event({ name: "ER" }), tasks: list }] } } }} />);
  const html = render(tasks);
  const table = html.slice(html.indexOf("wk-dp-tasks"));
  const cal = html.slice(html.indexOf("Detailed schedule"), html.indexOf("wk-dp-tasks"));

  it("shows the category of each task in its own column, — for a name of its own", () => {
    expect(table).toMatch(/<th class="fit">Done<\/th><th class="fit">Category<\/th><th class="l">Task<\/th><th class="fit">Start<\/th>/);
    expect([...table.matchAll(/<td class="l wk-dp-cat">(?:<span class="wk-dp-catv( sample)?">([^<]+)<\/span>|<span class="muted">(—)<\/span>)<\/td>/g)]
      .map((m) => (m[2] ?? m[3]) + (m[1] ? " (sample)" : ""))).toEqual(["Setup", "Sample making (sample)", "Sample shipment (sample)", "—"]);
  });

  it("turns the calendar row of a sample task into a count for each day, typed in place, its total by its name", () => {
    const qty = [...cal.matchAll(/aria-label="([^"]+) (\d{4}-\d\d-\d\d)" value="(\d+)"/g)].map((m) => `${m[1]} ${m[2]} ${m[3]}`);
    expect(qty).toEqual(["Sample making · 400pcs 2026-12-15 100", "Sample making · 400pcs 2026-12-16 300", "GN→AS 샘플발송 2026-12-16 50", "GN→AS 샘플발송 2026-12-19 300"]);
    // the calendar still fits the width (no scrolling), its count cells compact
    expect(cal).toContain('class="wk-gantt fill counts"');
    expect(cal).not.toContain("qopen");
    expect(cal).not.toMatch(/aria-label="Setup \d{4}/);
    expect(table).toContain('<span class="wk-dp-qtot" title="Samples of the days typed in the calendar above">Σ 400</span>');
  });

  it("charts the samples of each day above the calendar, made and shipped stacked with their sum", () => {
    expect(cal.indexOf('class="wk-chart"')).toBeGreaterThan(-1);
    expect(cal.indexOf('class="wk-chart"')).toBeLessThan(cal.indexOf("Sample making · 400pcs"));
    expect(cal).toMatch(/title="2026-12-16 · Sample making 300 · Sample shipment 50"><div class="col"><span class="v">350<\/span><i style="background:#eb6834;height:8px"><\/i><i style="background:#2a78d6;height:48px"><\/i>/);
    expect(cal).toMatch(/<i style="background:#2a78d6"><\/i>Sample making <b>400<\/b>/);
    expect(cal).toMatch(/<i style="background:#eb6834"><\/i>Sample shipment <b>350<\/b>/);
  });

  it("has no chart and no count cells without a sample task", () => {
    const none = render(tasks.filter((x) => !x.code?.startsWith("sample")));
    expect(none).not.toContain("wk-chart");
    expect(none).not.toContain("wk-qty-in");
  });

  it("keeps the counts with a sample task only", () => {
    const s = tasks[1];
    expect(keptTask(s, DEFAULT_VOCAB).task.qty).toEqual({ "2026-12-15": 100, "2026-12-16": 300 });
    expect("qty" in keptTask({ ...s, code: "setup" }, DEFAULT_VOCAB).task).toBe(false);
  });
});

describe("sample and material tasks of a list saved with codes of its own, sample sub-tasks counted day by day", () => {
  // the list as saved on the server (Setting › Dev Plan standard names): its own codes and Korean names
  const vocab: PlanVocab = {
    stages: DEFAULT_VOCAB.stages, templates: {},
    tasks: [
      { code: "SETUP LINE", name: "SETUP LINE", aliases: [] },
      { code: "Make Sample", name: "샘플 제작", aliases: [] },
      { code: "Send Sample", name: "샘플 발송", aliases: [] },
      { code: "Materials", name: "자재", aliases: [] },
    ],
  };
  const t = (id: string, name: string, f: Partial<PlanTask> = {}): PlanTask =>
    ({ id, name, start: "2026-12-03", end: "2026-12-10", owner: "", done: false, note: "", ...f });
  const tasks: PlanTask[] = [
    t("m", "샘플 제작 · 800SET", { code: "Make Sample", detail: "800SET" }),
    t("m1", "1차 400SET", { parentId: "m", qty: { "2026-12-04": 150, "2026-12-05": 250 } }),
    t("m2", "2차 400SET", { parentId: "m", qty: { "2026-12-05": 100 } }),
    t("m3", "자재 확인", { parentId: "m", code: "Materials", qty: { "2026-12-05": 999 } }),
    t("s", "샘플 발송 · 500SET", { code: "Send Sample", detail: "500SET", qty: { "2026-12-10": 500 } }),
    t("x", "자재", { code: "Materials" }),
  ];
  const html = renderToStaticMarkup(<DevPlan onDirty={() => {}} onSaved={() => {}} onError={() => {}} vocab={vocab}
    model={{ name: "M", revision: 1, updatedAt: "", updatedBy: "", info: { devPlan: { events: [{ ...event({ name: "ER" }), tasks }] } } }} />);
  const table = html.slice(html.indexOf("wk-dp-tasks"));
  const cal = html.slice(html.indexOf("Detailed schedule"), html.indexOf("wk-dp-tasks"));
  const typed = [...cal.matchAll(/aria-label="([^"]+) (\d{4}-\d\d-\d\d)" value="(\d+)"/g)].map((m) => `${m[1]} ${m[2]} ${m[3]}`);

  it("knows the sample tasks by their names (Make Sample, 샘플 제작), so their days are counted and charted", () => {
    expect(typed).toContain("샘플 발송 · 500SET 2026-12-10 500");
    expect(cal).toContain('class="wk-chart"');
    expect(cal).toMatch(/<i style="background:#2a78d6"><\/i>샘플 제작 <b>500<\/b>/);
    expect(cal).toMatch(/<i style="background:#eb6834"><\/i>샘플 발송 <b>500<\/b>/);
  });

  it("gives a sub-task of a sample task (no category of its own) its count cells, the task the sum of them", () => {
    expect(typed).toEqual(expect.arrayContaining(["1차 400SET 2026-12-04 150", "1차 400SET 2026-12-05 250", "2차 400SET 2026-12-05 100"]));
    // the task itself: the sum of its sample sub-tasks each day, shown only
    expect(typed.some((x) => x.startsWith("샘플 제작 · 800SET"))).toBe(false);
    expect(cal).toContain('<td class="q sum in has" style="box-shadow:inset 0 -3px 0 #');
    expect([...cal.matchAll(/<td class="q sum[^"]*"[^>]*>(\d[\d,]*)<\/td>/g)].map((m) => m[1])).toEqual(["150", "350"]);
    // a sub-task with a category of its own that is no sample task is not counted
    expect(cal).not.toMatch(/aria-label="자재 확인 \d{4}/);
    expect(table).toContain("Σ 500</span>");
    expect(cal).toMatch(/title="2026-12-05 · 샘플 제작 350">/);
  });

  it("knows the material task by its name (Materials, 자재): its row has the button to pick materials", () => {
    expect(table).toContain("+ Materials");
  });

  it("keeps the counts of a sub-task of a sample task", () => {
    expect(keptTask(tasks[1], vocab, tasks).task.qty).toEqual({ "2026-12-04": 150, "2026-12-05": 250 });
    expect("qty" in keptTask(tasks[3], vocab, tasks).task).toBe(false);
  });
});
