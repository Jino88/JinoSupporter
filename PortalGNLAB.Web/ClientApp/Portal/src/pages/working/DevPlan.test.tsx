import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { addDays } from "./dates";
import DevPlan, { EventEdit, Gantt } from "./DevPlan";
import type { PlanEvent } from "./plan";

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

  it("puts the sub-tasks on the calendar under their task, the task over their days when it has none", () => {
    const cal = html.slice(html.indexOf("Detailed schedule"), html.indexOf("wk-dp-tasks"));
    expect([...cal.matchAll(/<td class="nm( child)?">[^]*?<b class="">([^<]+)<\/b><small>([^<]*)<\/small>/g)].map((m) => `${m[2]}${m[1] ?? ""} ${m[3]}`))
      .toEqual(["SETUP 11-30 – 12-02", "Line A child 11-30 – 12-01", "Line B child 12-02", "GN parts 11-27"]);
  });
});
