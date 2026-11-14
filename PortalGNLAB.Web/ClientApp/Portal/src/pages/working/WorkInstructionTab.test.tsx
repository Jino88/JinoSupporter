import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { InstructionBody } from "./WorkInstructionTab";
import type { InstructionDoc } from "./workInstruction";

const doc: InstructionDoc = {
  fileName: "MSM-G778B.xlsx", uploadedAt: "2026-11-14T03:05:00", uploadedBy: "kim", size: 2048,
  revisions: [{ date: "26-May", page: "49", rev: "0", content: "Ban hành", author: "Mr.Dung" }],
  processes: [
    { sheet: "Xếp UPPER", code: "B1", name: "Xếp UPPER", standardTime: 12.312, spareTime: 1.231, page: "1/53", ctq: false,
      parts: [{ no: "1", name: "ENCLOSURE", spec: "ENC-G778B", qty: "1 PCS" }],
      tools: [{ no: "1", name: "JIG", note: "ENCLOSURE DISPENSOR JIG (V4)" }],
      steps: [{ no: 1, text: "Chuẩn bị UPPER." }], keyPoints: ["1.Không có dị vật."] },
    { sheet: "Kiểm tra Bako", code: "C15", name: "Kiểm tra Bako", standardTime: 18.468, spareTime: 1.85, page: "44/53", ctq: true,
      parts: [], tools: [], steps: [], keyPoints: [] },
  ],
};
const body = (d: InstructionDoc | null, open: string | null = null, busy = false) =>
  renderToStaticMarkup(<InstructionBody doc={d} busy={busy} open={open} onOpen={() => {}} onPick={() => {}} onDownload={() => {}} onDelete={() => {}} />);

describe("Work Instructions tab", () => {
  it("asks for the workbook when none is uploaded", () => {
    const html = body(null);
    expect(html).toContain("Upload workbook");
    expect(html).not.toContain("Replace");
    expect(html).not.toContain("<table");
  });

  it("shows the file, Replace / Download / Delete, the revision history and the process table", () => {
    const html = body(doc);
    for (const t of ["MSM-G778B.xlsx", "kim", "2 KB", ">Replace<", ">Download<", ">Delete<", "Revision history", "Mr.Dung", "Ban hành"]) expect(html).toContain(t);
    expect(html).not.toContain("Upload workbook");
    expect(html).toContain("Processes (2)");
    expect(html).toMatch(/<td>1<\/td><td>B1<\/td><td class="l">Xếp UPPER<\/td><td>12.312<\/td><td class="l">ENCLOSURE DISPENSOR JIG \(V4\)<\/td><td><\/td>/);
    expect(html).toMatch(/<td>C15<\/td>.*<td><span class="wi-ctq">CTQ<\/span><\/td>/);
    expect(html).not.toContain("wi-detail");
  });

  it("opens the detail under the clicked row only: parts, equipment & tools, steps, key control items", () => {
    const html = body(doc, "Xếp UPPER");
    expect(html.split("wi-detail").length - 1).toBe(1);
    for (const t of ["ENC-G778B", "1 PCS", "Equipment &amp; tools", "Chuẩn bị UPPER.", "1.Không có dị vật.", "Key control items"]) expect(html).toContain(t);
    expect(html.indexOf("wi-detail")).toBeGreaterThan(html.indexOf("<td>B1</td>"));
    expect(html.indexOf("wi-detail")).toBeLessThan(html.indexOf("<td>C15</td>"));
  });

  it("disables the buttons while the workbook is being read", () => {
    expect(body(doc, null, true)).toMatch(/<button class="btn sm" disabled=""[^>]*>Reading…<\/button>/);
  });
});
