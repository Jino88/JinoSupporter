import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { IssueForm } from "./IssueTable";
import { blankEntry } from "./logic";

describe("Issue form", () => {
  it("has four columns: classification, owner and dates | issue and detail | files | final result", () => {
    const html = renderToStaticMarkup(
      <IssueForm value={{ ...blankEntry("issue", "2026-11-14", "MP"), title: "RA2/RM MOVING" }} images={[]}
        lists={{ types: [], processes: [], causes: [] }} busy={false} onSave={() => {}} onCancel={() => {}} onOpen={() => {}} />,
    );
    const columns = html.split('<div class="wk-if-col').slice(1)
      .map((part) => [...part.matchAll(/<label>([^<]+)<\/label>/g)].map((m) => m[1]));
    expect(columns).toEqual([
      ["Dev phase", "Issue type", "Process", "Cause group", "Owner", "Registered", "Due", "Done"],
      ["Issue", "Issue detail"],
      ["Files"],
      ["Final result"],
    ]);
  });
});
