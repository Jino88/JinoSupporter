import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { EventEdit } from "./DevPlan";
import type { PlanEvent } from "./plan";

const event = (f: Partial<PlanEvent>): PlanEvent =>
  ({ id: "e", name: "DR build", start: "", end: "", color: "#2563eb", tasks: [], rows: [], bomFrom: "", ...f });

const years = (html: string) => [...html.matchAll(/<input class="inp sm yr"[^>]*aria-label="(Start|End) year" value="(\d*)"/g)].map((m) => `${m[1]} ${m[2]}`);

describe("event dates as yyyy/mm/dd", () => {
  it("shows the event Start and End as a year box and the month/day typed, not the browser's date box", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({ start: "2027-01-05", end: "2028-02-28" })} onKeep={() => {}} onCancel={() => {}} />);
    expect(html).not.toContain('type="date"');
    expect(years(html)).toEqual(["Start 2027", "End 2028"]);
    expect(html).toContain('value="01/05"');
    expect(html).toContain('value="02/28"');
    expect(html).toContain('placeholder="mm/dd"');
  });

  it("puts this year in the year box of empty event dates", () => {
    const html = renderToStaticMarkup(<EventEdit value={event({})} onKeep={() => {}} onCancel={() => {}} />);
    const y = String(new Date().getFullYear());
    expect(years(html)).toEqual([`Start ${y}`, `End ${y}`]);
  });
});
