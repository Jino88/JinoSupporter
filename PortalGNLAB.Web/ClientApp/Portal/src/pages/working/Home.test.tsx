import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import Home from "./Home";
import type { ModelLoad } from "./homeSummary";
import type { ModelInfo } from "./types";

const load = (name: string, info: ModelInfo): ModelLoad =>
  ({ name, path: ["MODULE", "MSM"], model: { name, info, revision: 1, updatedAt: "2026-11-10T09:00:00", updatedBy: "a" }, entries: [], error: "" });
const event = (name: string, start: string, end: string) => ({ id: name, name, start, end, color: "#4f46e5", tasks: [], rows: [], bomFrom: "" });

describe("Working Home", () => {
  it("shows the development plans first, then every model", () => {
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[
      load("MSM-A386B", { phase: "PV", customer: "Acme", devPlan: { events: [event("PV build", "2026-11-01", "2026-11-20"), event("MP", "2026-12-01", "2026-12-05")] } }),
      load("MSM-OLD", { phase: "MP", devPlan: { events: [event("Old", "2026-11-01", "2026-11-02")] } }),
      load("MSM-NEW", { phase: "ER" }),
    ]} />);
    expect(html.indexOf("Development plans")).toBeGreaterThan(-1);
    expect(html.indexOf("Development plans")).toBeLessThan(html.indexOf(">Models<"));
    const plans = html.slice(0, html.indexOf(">Models<"));
    expect(plans).toContain(">MSM-A386B</button>");
    expect(plans).not.toContain(">MSM-OLD</button>");
    expect(plans).toContain("Now: <b>PV build</b>");
    expect(plans).toContain("Next: <b>MP 2026/12/01</b>");
    expect(plans).toContain("2026/11/01 – 11/20");
    expect(plans).toMatch(/In development without a plan:.*>MSM-NEW<\/button>/);
    const table = html.slice(html.indexOf(">Models<"));
    expect([...table.matchAll(/class="wk-hm-link">([^<]+)</g)].map((m) => m[1])).toEqual(["MSM-A386B", "MSM-OLD", "MSM-NEW"]);
    expect(table).toContain("<td>Acme</td>");
  });

  it("says no model in development has a plan yet", () => {
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[load("A", { phase: "DR" })]} />);
    expect(html).toContain("No model in development has a Dev Plan yet");
  });
});
