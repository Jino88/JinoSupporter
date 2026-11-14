import type { ReactNode } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { Gantt } from "./DevPlan";
import Home from "./Home";
import { TrContext, lookUp, writtenIn } from "./i18n";
import IssueTable from "./IssueTable";
import type { Entry } from "./types";

const ko: Record<string, string> = {
  "Development plans": "개발 계획", "Material in": "자재 입고", "Jig ready": "지그 준비", "Buzz noise": "잡음 발생",
  "Fixed the glue": "접착제 수정", "Issue": "이슈",
};
const tr = (t: string) => lookUp(ko, t);
const inKorean = (node: ReactNode) => renderToStaticMarkup(<TrContext.Provider value={tr}>{node}</TrContext.Provider>);

describe("Working shown in another language", () => {
  it("shows a Korean text as written in Korean, not through a translation that can reword it", () => {
    expect(writtenIn("ko", "REAR AP THF 본드로 인한 변형")).toBe(true);
    expect(writtenIn("ko", "ㅇㅋ")).toBe(true);
    expect(writtenIn("ko", "Jig ready")).toBe(false);
    expect(writtenIn("en", "본드로 인한 변형")).toBe(false);
    expect(writtenIn("", "본드")).toBe(false);
  });

  it("looks a text up by its trimmed text and keeps it as written without a translation", () => {
    expect(lookUp(ko, " Jig ready ")).toBe("지그 준비");
    expect(lookUp(ko, "MSM-G778B")).toBe("MSM-G778B");
    expect(lookUp(ko, "")).toBe("");
  });

  it("shows Home's plan headings, events and tasks translated, model names as written", () => {
    const task = (id: string, name: string) => ({ id, name, start: "2026-11-18", end: "2026-11-18", owner: "", done: false, note: "" });
    const html = inKorean(<Home models={[]} today="2026-11-11" onOpen={() => {}} initial={[{
      name: "MSM-G778B", path: ["MODULE"], entries: [], error: "",
      model: { name: "MSM-G778B", revision: 1, updatedAt: "", updatedBy: "", info: { phase: "PV", devPlan: { events: [{
        id: "e", name: "LPR", start: "2026-11-17", end: "2026-11-20", color: "#4f46e5", rows: [], bomFrom: "", tasks: [task("a", "Material in"), task("b", "Jig ready")],
      }] } } },
    }]} />);
    expect(html).toContain("개발 계획");
    expect(html).toContain(">자재 입고<");
    expect(html).toContain(">지그 준비<");
    expect(html).toContain(">MSM-G778B<");
    expect(html).not.toContain(">Material in<");
  });

  it("shows Dev Plan calendar rows and issue titles and notes translated", () => {
    const gantt = inKorean(<Gantt rows={[{ id: "a", label: "Material in", sub: "", span: null, color: "#000" }]} days={["2026-11-18"]} />);
    expect(gantt).toContain(">자재 입고<");
    const issue = {
      id: "i", kind: "issue", seq: 1, date: "2026-11-10", title: "Buzz noise", body: "", status: "Open", severity: "", result: "", owner: "",
      due: "", phase: "DR", category: "", process: "", cause: "", doneDate: "", finalResult: "Fixed the glue", logs: [], images: [],
      modelName: "M", revision: 1, createdBy: "", createdAt: "", updatedBy: "", updatedAt: "",
    } as Entry;
    const html = inKorean(<IssueTable modelName="M" modelPhase="DR" entries={[issue]} today="2026-11-11" onChanged={() => {}} onError={() => {}} />);
    expect(html).toContain("<b>잡음 발생</b>");
    expect(html).toContain(">접착제 수정<");
    expect(html).toContain(">이슈<");
  });
});
