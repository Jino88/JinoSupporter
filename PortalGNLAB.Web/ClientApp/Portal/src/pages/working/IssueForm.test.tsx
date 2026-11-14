import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { IssueForm } from "./IssueTable";
import { blankEntry } from "./logic";

describe("Issue form", () => {
  it("puts owner and dates in one left column and issue, detail and files side by side on the right", () => {
    const html = renderToStaticMarkup(
      <IssueForm value={{ ...blankEntry("issue", "2026-11-14", "MP"), title: "RA2/RM MOVING" }} images={[]}
        lists={{ types: [], processes: [], causes: [] }} busy={false} onSave={() => {}} onCancel={() => {}} onOpen={() => {}} />,
    );
    const labels = (part: string) => [...part.matchAll(/<label>([^<]+)<\/label>/g)].map((m) => m[1]);
    const side = html.slice(html.indexOf('class="wk-if-side"'), html.indexOf('class="wk-if-body"'));
    const body = html.slice(html.indexOf('class="wk-if-body"'));
    expect(labels(side)).toEqual(["Owner", "Registered", "Due", "Done"]);
    expect(labels(body).slice(0, 4)).toEqual(["Issue", "Issue detail", "Files", "Final result"]);
    expect(html).not.toContain("Click here and press Ctrl+V");
  });
});
