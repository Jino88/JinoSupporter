import { renderToStaticMarkup } from "react-dom/server";
import { PersonalPage } from "./WorkingPersonal";
import { personalKey } from "./working/personal";

describe("Working › Personal", () => {
  it("keeps a user's plans under a model of their own name", () => {
    expect(personalKey(" alice ")).toBe("~personal:alice");
    expect(personalKey("bob")).not.toBe(personalKey("alice"));
  });

  it("shows the user's plans on the Issues register", () => {
    const issue = (id: string, seq: number, title: string, x: object = {}) => ({ id, seq, title, kind: "issue", status: "Open", doneDate: "", due: "", phase: "",
      body: "", owner: "", date: "2026-11-10", category: "", process: "", cause: "", finalResult: "", logs: [], images: [], createdAt: "2026-11-10", ...x });
    const html = renderToStaticMarkup(<PersonalPage user="alice" lang="" waiting={0} onLang={() => {}} initial={[
      issue("1", 1, "Write the DV report"), issue("2", 2, "Book the jig", { doneDate: "2026-11-11" }), { ...issue("3", 1, "a dev note"), kind: "dev" },
    ] as never} />);
    expect(html).toContain("Working · Personal");
    expect(html).toContain("Only you can see these plans.");
    expect(html).toContain("My Plans");
    expect(html).toContain("Write the DV report");
    expect(html).not.toContain("Book the jig"); // closed, the Open filter is shown first
    expect(html).not.toContain("a dev note");
    expect(html).toContain("+ New issue");
  });
});
