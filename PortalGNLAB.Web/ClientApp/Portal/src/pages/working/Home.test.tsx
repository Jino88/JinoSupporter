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
    expect(plans).toContain("<i>Now</i>PV build</span>");
    expect(plans).toContain("<i>Next</i>MP · 2026/12/01</span>");
    expect(plans).toContain("2026/11/01 – 11/20");
    expect(plans).toMatch(/In development without a plan:.*>MSM-NEW<\/button>/);
    const table = html.slice(html.indexOf(">Models<"));
    expect([...table.matchAll(/class="wk-hm-link">([^<]+)</g)].map((m) => m[1])).toEqual(["MSM-A386B", "MSM-OLD", "MSM-NEW"]);
    expect(table).toContain("<td>Acme</td>");
  });

  it("writes the days of a short event beside its bar, not cut inside it", () => {
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[
      load("MSM-G778B", { phase: "PV", devPlan: { events: [event("LPR", "2026-11-18", "2026-11-20"), event("MP", "2027-01-20", "2027-01-20"), event("Long", "2026-11-12", "2027-01-10")] } }),
    ]} />);
    expect(html).toMatch(/class="wk-hm-blabel" style="left:calc\([\d.]+% \+ 6px\)"[^>]*>2026\/11\/18 – 11\/20</);
    expect(html).toMatch(/class="wk-hm-blabel" style="right:calc\([\d.]+% \+ 6px\)"[^>]*>2027\/01\/20</);
    // a long bar too: its days beside it, never inside
    expect(html).toMatch(/class="wk-hm-bar upcoming"[^>]*><\/span><span class="wk-hm-blabel"[^>]*>2026\/11\/12 – 01\/10</);
    expect(html).not.toContain("wk-hm-btext");
    expect(html).not.toMatch(/class="wk-hm-bar[^"]*"[^>]*>2026\/11\/18/);
  });

  it("shows every task of an event under it: name, owner, done or late, and its days beside its bar", () => {
    const task = (id: string, name: string, start: string, end: string, done = false, owner = "") => ({ id, name, start, end, owner, done, note: "" });
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[
      load("MSM-G778B", { phase: "PV", devPlan: { events: [{ ...event("LPR", "2026-11-18", "2026-11-21"), tasks: [
        task("a", "Material in", "2026-11-05", "2026-11-09", true, "Kim"), task("b", "Jig ready", "2026-11-02", "2026-11-06", false, "Park"),
        task("c", "Assy & test", "2026-11-20", "2026-12-30"), task("d", "Report", "", ""),
      ] }] } }),
    ]} />);
    const rows = [...html.matchAll(/<div class="wk-hm-row task (\w+)">.*?class="wk-hm-task" title="[^"]*">([^<]+)<\/span>(?:<span class="wk-hm-owner"[^>]*>([^<]+)<)?/g)]
      .map((m) => [m[2], m[1], m[3] ?? ""]);
    expect(rows).toEqual([["Material in", "done", "Kim"], ["Jig ready", "late", "Park"], ["Assy &amp; test", "open", ""], ["Report", "open", ""]]);
    expect(html).toMatch(/class="wk-hm-tbar done"[^>]*><\/span><span class="wk-hm-blabel"[^>]*>11\/05 – 11\/09</);
    expect(html).toMatch(/class="wk-hm-tbar open"[^>]*><\/span><span class="wk-hm-blabel"[^>]*>11\/20 – 12\/30</);
    expect(html).toContain('<span class="wk-hm-st undated">No date</span>');
    expect(html).toContain('<i>Tasks</i>1/4');
    expect(html).toContain("1 late task<");
    // the event bar shows its days only, its task count is in the label column
    expect(html).toMatch(/class="wk-hm-bar [a-z]+"[^>]*>.*?<\/span><span class="wk-hm-blabel"[^>]*>2026\/11\/18 – 11\/21<\/span>/);
  });

  it("says no model in development has a plan yet", () => {
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[load("A", { phase: "DR" })]} />);
    expect(html).toContain("No model in development has a Dev Plan yet");
  });

  it("shows the MP plans (Setup Plans by line and process) apart from the development plans", () => {
    const setup = [{ id: "r1", no: 1, line: "L1", processName: "Bonding", machinePlan: "2026-11-01", setupPlan: "2026-11-03", setupStatus: "O", setupFinish: "",
      layoutNo: "", machineState: "", machineNote: "", modelName: "MSM-MP", revision: 1, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "" }] as never;
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[
      load("MSM-A386B", { phase: "PV", devPlan: { events: [event("PV build", "2026-11-01", "2026-11-20")] } }),
      { ...load("MSM-MP", { phase: "MP" }), setup },
    ]} />);
    const dev = html.slice(html.indexOf("Development plans"), html.indexOf(">MP plans<"));
    const mp = html.slice(html.indexOf(">MP plans<"), html.indexOf(">Models<"));
    expect(html.indexOf(">MP plans<")).toBeGreaterThan(html.indexOf("Development plans"));
    expect(dev).toContain("PV build");
    expect(dev).not.toContain("MSM-MP");
    expect(mp).toContain(">MSM-MP<");
    expect(mp).toContain("Model · line · process");
    expect(mp).toContain(">Bonding<");
    expect(mp).toContain("<i>Setup done</i>1/1");
    expect(mp).not.toContain("PV build");
  });
});
