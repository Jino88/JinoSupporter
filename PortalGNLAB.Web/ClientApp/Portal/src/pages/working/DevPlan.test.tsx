import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { EventEdit } from "./DevPlan";
import type { PlanEvent } from "./plan";

const event = (f: Partial<PlanEvent>): PlanEvent =>
  ({ id: "e", name: "DR build", start: "", end: "", color: "#2563eb", tasks: [], rows: [], bomFrom: "", ...f });

describe("event dates as yyyy/mm/dd", () => {
  it("shows the event Start and End as the fixed year and the month/day typed, not the browser's date box", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({ start: "2027-01-05", end: "2027-02-28" })} onKeep={() => {}} onCancel={() => {}} />);
    expect(html).not.toContain('type="date"');
    expect(html.match(/<span class="y">2027\/<\/span>/g)).toHaveLength(2);
    expect(html).toContain('value="01/05"');
    expect(html).toContain('value="02/28"');
    expect(html).toContain('placeholder="mm/dd"');
  });

  it("shows this year beside empty event dates", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({})} onKeep={() => {}} onCancel={() => {}} />);
    expect(html.match(/<span class="y">(\d{4})\/<\/span>/)?.[1]).toBe(String(new Date().getFullYear()));
  });
});
