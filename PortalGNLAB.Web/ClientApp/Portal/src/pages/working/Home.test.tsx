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

  it("folds the tasks of an event away under its arrow, the others staying open", () => {
    const task = (id: string, name: string) => ({ id, name, start: "2026-11-19", end: "2026-11-19", owner: "", done: false, note: "" });
    const loads = [load("MSM-G778B", { phase: "PV", devPlan: { events: [
      { ...event("LPR", "2026-11-17", "2026-11-20"), tasks: [task("a", "SEVT AUDIT")] },
      { ...event("PV", "2026-11-30", "2026-12-11"), tasks: [task("b", "SETUP"), task("c", "R&D Business Trip")] },
      event("MP", "2026-12-07", "2026-12-07"),
    ] } })];
    const open = renderToStaticMarkup(<Home models={[]} today="2026-11-17" onOpen={() => {}} initial={loads} folded={[]} />);
    expect(open).toContain('aria-expanded="true" aria-label="Fold the tasks of LPR"');
    expect(open).toContain('aria-label="Fold the tasks of PV"');
    // an event with no task has nothing to fold
    expect(open).not.toContain("tasks of MP");
    expect(open).toContain(">SEVT AUDIT<");
    expect(open).toContain(">SETUP<");

    const shut = renderToStaticMarkup(<Home models={[]} today="2026-11-17" onOpen={() => {}} initial={loads} folded={["dev|MSM-G778B|PV"]} />);
    expect(shut).toContain('class="wk-hm-fold shut" aria-expanded="false" aria-label="Unfold the tasks of PV"');
    expect(shut).not.toContain(">SETUP<");
    expect(shut).not.toContain(">R&amp;D Business Trip<");
    // the event row itself and its count stay; LPR keeps its task
    expect(shut).toContain('<span class="wk-hm-count">0/2</span>');
    expect(shut).toContain(">SEVT AUDIT<");
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

  it("shows the MP issues of every model as their own page, the open ones until All is picked", () => {
    const issue = (id: string, seq: number, title: string, x: object = {}) => ({ id, seq, title, kind: "issue", status: "Open", doneDate: "", due: "", phase: "MP",
      body: "", owner: "Kim", date: "2026-11-02", category: "Quality", process: "SPK Assy", cause: "", finalResult: "", logs: [], ...x });
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" view="mpissues" onOpen={() => {}} initial={[
      { ...load("MSM-MP", { phase: "MP" }), entries: [issue("1", 1, "AOI false call", { due: "2026-11-05" }), issue("2", 2, "Closed one", { doneDate: "2026-11-03" }), issue("3", 3, "Dev one", { phase: "DV" })] as never },
    ]} />);
    expect(html).toContain(">MP Issues<");
    expect(html).not.toContain("Development plans");
    expect(html).toContain("<b>AOI false call</b>");
    expect(html).toMatch(/class="late">26\/11\/05<span class="wk-late">Late<\/span>/);
    expect(html).not.toContain("Closed one");
    expect(html).not.toContain("Dev one");
    expect(html).toContain(">Open 1</button><button class=\"wk-chip\">All 2</button>");
  });

  it("shows the Dev issues of every model as their own page, opening a model on its Dev Issue tab", () => {
    const issue = (id: string, seq: number, title: string, x: object = {}) => ({ id, seq, title, kind: "issue", status: "Open", doneDate: "", due: "", phase: "DV",
      body: "", owner: "Kim", date: "2026-11-02", category: "Quality", process: "SPK Assy", cause: "", finalResult: "", logs: [], ...x });
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" view="devissues" onOpen={() => {}} initial={[
      { ...load("MSM-DV", { phase: "DV" }), entries: [issue("1", 1, "Speaker buzz"), issue("2", 2, "MP one", { phase: "MP" }), issue("3", 3, "Setup one", { setupRowId: "r1" })] as never },
    ]} />);
    expect(html).toContain(">Dev Issues<");
    expect(html).not.toContain(">MP Issues<");
    expect(html).toContain("<b>Speaker buzz</b>");
    expect(html).not.toContain("MP one");
    expect(html).not.toContain("Setup one");
    expect(html).toContain("title=\"Open the model&#x27;s Dev Issue\">MSM-DV</button>");
    expect(html).toContain(">Open 1</button><button class=\"wk-chip\">All 1</button>");
  });

  it("shows under an MP issue its notes not done and its open sub-issues", () => {
    const note = (id: string, text: string, x: object = {}) => ({ id, text, date: "2026-11-03", createdAt: "2026-11-03", owner: "Lee", due: "", doneDate: "", createdBy: "a", ...x });
    const issue = (id: string, seq: number, title: string, x: object = {}) => ({ id, seq, title, kind: "issue", status: "Open", doneDate: "", due: "", phase: "MP",
      body: "", owner: "Kim", date: "2026-11-02", category: "ETC", process: "RA2", cause: "", finalResult: "", logs: [], ...x });
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" view="mpissues" onOpen={() => {}} initial={[
      { ...load("TIM-338", { phase: "MP" }), entries: [
        issue("p", 1, "RA2/RM MOVING", { logs: [note("n1", "Check the jig"), note("n2", "Old step", { doneDate: "2026-11-04" })] }),
        issue("a", 2, "RA2 ISSUE", { parentId: "p", logs: [note("n3", "Wait for parts", { due: "2026-11-05" })] }),
        issue("b", 3, "RM ISSUE", { parentId: "p", doneDate: "2026-11-06" }),
      ] as never },
    ]} />);
    expect(html).toContain("↳ Check the jig");
    expect(html).not.toContain("Old step");
    expect(html).toContain("↳ <b>RA2 ISSUE</b>");
    expect(html).toMatch(/1\.1-1<\/td>.*↳ Wait for parts.*class="late">26\/11\/05<span class="wk-late">Late/);
    expect(html).not.toContain("RM ISSUE");
    expect(html).toContain(">Open 2</button><button class=\"wk-chip\">All 3</button>");
  });

  it("shows the issue on one line and the rest in the Issue detail column", () => {
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-11" view="mpissues" onOpen={() => {}} initial={[
      { ...load("TIM-338", { phase: "MP" }), entries: [{
        id: "f", seq: 1, title: "FINAL LINE ISSUE", kind: "issue", status: "Open", doneDate: "", due: "", phase: "MP", body: "Line A\nLine B",
        owner: "Kim", date: "2026-11-02", category: "ETC", process: "FINAL", cause: "", finalResult: "",
        logs: [
          { id: "n1", text: "QM stock shared", detail: "Order qty to discuss", date: "2026-11-03", createdAt: "2026-11-03", owner: "", due: "", doneDate: "", createdBy: "a" },
          { id: "n2", text: "Old note\nsecond line", date: "2026-11-04", createdAt: "2026-11-04", owner: "", due: "", doneDate: "", createdBy: "a" },
        ],
      }] as never },
    ]} />);
    expect(html).toContain("<th>Issue</th><th>Issue detail</th><th>Owner</th>");
    expect(html).toContain("<b>FINAL LINE ISSUE</b></td><td class=\"l detail\">Line A\nLine B</td>");
    expect(html).toContain("↳ QM stock shared</span></td><td class=\"l detail\">Order qty to discuss</td>");
    expect(html).toContain("↳ Old note</span></td><td class=\"l detail\">second line</td>");
  });

  it("draws the actual days under the planned ones: green on plan, red after the planned end, still going up to today", () => {
    const task = (id: string, name: string, start: string, end: string, f: Record<string, unknown> = {}) => ({ id, name, start, end, owner: "", done: false, note: "", ...f });
    const html = renderToStaticMarkup(<Home models={[]} today="2026-11-16" onOpen={() => {}} initial={[
      load("MSM-G778B", { phase: "PV", devPlan: { events: [{ ...event("LPR", "2026-11-10", "2026-11-20"), tasks: [
        task("a", "Jig ready", "2026-11-10", "2026-11-12", { done: true, actualStart: "2026-11-10", actualEnd: "2026-11-12" }),
        task("b", "Material in", "2026-11-11", "2026-11-13", { done: true, actualStart: "2026-11-12", actualEnd: "2026-11-15" }),
        task("c", "Assy", "2026-11-14", "2026-11-18", { actualStart: "2026-11-14" }),
        task("d", "Report", "2026-11-19", "2026-11-20"),
      ] }] } }),
    ]} />);
    expect(html).toMatch(/class="wk-hm-abar"[^>]*title="Jig ready: 11\/10 – 11\/12 · done · Act. 11\/10 – 11\/12"/);
    expect(html).toMatch(/class="wk-hm-abar over"[^>]*title="Material in[^"]*Act. 11\/12 – 11\/15"/);
    expect(html).toMatch(/class="wk-hm-abar going"[^>]*title="Assy[^"]*Act. 11\/14 –"/);
    expect(html).toContain('11/11 – 11/13 · <span class="wk-hm-act over">Act. 11/12 – 11/15</span>');
    expect(html).toContain('11/10 – 11/12 · <span class="wk-hm-act">Act. 11/10 – 11/12</span>');
    // no actual days, no line
    expect(html.match(/class="wk-hm-abar/g)).toHaveLength(4); // 3 tasks + the event
    // the event: from the first actual start of its tasks, still going while one is
    expect(html).toContain('2026/11/10 – 11/20 · <span class="wk-hm-act">Act. 11/10 –</span>');
  });
});
