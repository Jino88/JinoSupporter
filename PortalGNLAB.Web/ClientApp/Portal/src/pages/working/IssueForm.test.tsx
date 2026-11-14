import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { IssueForm, LogEditRow } from "./IssueTable";
import { blankEntry, blankLog } from "./logic";

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

  it("edits a note in a panel across the table: owner and dates | issue and detail | files | buttons", () => {
    const html = renderToStaticMarkup(
      <table><tbody>
        <LogEditRow no="5-1" value={{ ...blankLog("2026-11-14"), text: "1. BN 불량 현황", detail: "line 1\nline 2" }} images={[]} today="2026-11-14"
          busy={false} isNew={false} onSave={async () => true} onCancel={() => {}} onOpen={() => {}} />
      </tbody></table>,
    );
    expect(html).toContain('<td colSpan="13" class="wk-ne-cell">');
    const part = (cls: string) => html.slice(html.indexOf(`class="${cls}"`));
    expect([...part("wk-ne-fields").split('class="wk-ne-text"')[0].matchAll(/<label>([A-Za-z]+)/g)].map((m) => m[1]))
      .toEqual(["Owner", "Date", "Due", "Done"]);
    expect(part("wk-ne-text")).toMatch(/^[^]*?value="1. BN 불량 현황"[^]*?<textarea[^>]*>line 1\nline 2<\/textarea>/);
    expect(html.indexOf('class="wk-ne-text"')).toBeLessThan(html.indexOf("wk-ne-files"));
    expect(part("wk-ne-act")).toMatch(/Cancel<\/button><button[^>]*>Save<\/button>/);
  });
});
