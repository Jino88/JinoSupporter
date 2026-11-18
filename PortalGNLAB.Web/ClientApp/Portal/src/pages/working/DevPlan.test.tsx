import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { addDays } from "./dates";
import DevPlan, { afterSave, autoSaveDue, editGroup, EventEdit, Gantt, keptTask, openRows, TaskEdit, TaskEditRow, TidyNames, withRows } from "./DevPlan";
import { tidyRows } from "./planNames";
import { DEFAULT_VOCAB } from "./planVocab";
import type { DevPlan as Plan, PlanEvent } from "./plan";

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
    expect(sel(std, "Standard task")).toContain('<option value="sample-send" selected="">Sample shipment</option>');
    expect(std).toMatch(/aria-label="Detail" value="300ea"/);
    const own = renderToStaticMarkup(<table><tbody><TaskEdit value={{ ...t, name: "Report", code: undefined, detail: undefined }} vocab={DEFAULT_VOCAB} onKeep={() => {}} onCancel={() => {}} /></tbody></table>);
    expect(sel(own, "Standard task")).toContain('<option value="" selected="">Other (own name)</option>');
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
