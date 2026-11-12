import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { addDays } from "../productionInput/logic";
import { EventEdit, Gantt } from "./DevPlan";
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
