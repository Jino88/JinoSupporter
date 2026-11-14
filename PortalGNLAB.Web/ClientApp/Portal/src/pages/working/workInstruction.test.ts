import { describe, expect, it } from "vitest";
import { equipmentText, sizeText, timeText, whenText, type InstructionProcess } from "./workInstruction";

const process = (tools: { name: string; note: string }[]): InstructionProcess => ({
  sheet: "S", code: "B1", name: "Xếp UPPER", standardTime: 12.312, spareTime: 1.231, page: "1/53", ctq: false, parts: [], steps: [], keyPoints: [],
  tools: tools.map((t, i) => ({ no: String(i + 1), ...t })),
});

describe("Work Instructions helpers", () => {
  it("names the machine and the jig of a process by the notes of their rows, once each", () => {
    const p = process([
      { name: "IONIZER", note: "SD-95" },
      { name: "Máy kiểm tra", note: "BAKO 2130C" },
      { name: "JIG", note: "SUB-BAFFLE JIG (V4)" },
      { name: "jig", note: "SUB-BAFFLE JIG (V4)" },
      { name: "Bao ngón", note: "Thay 2h/lần" },
      { name: "장비", note: "" },
    ]);
    expect(equipmentText(p)).toBe("BAKO 2130C / SUB-BAFFLE JIG (V4) / 장비");
    expect(equipmentText(process([{ name: "IONIZER", note: "SD-95" }]))).toBe("");
  });

  it("writes a time with up to 3 decimals and nothing when there is none", () => {
    expect(timeText(12.312)).toBe("12.312");
    expect(timeText(6.1)).toBe("6.1");
    expect(timeText(18)).toBe("18");
    expect(timeText(null)).toBe("");
  });

  it("writes sizes and times for the file line", () => {
    expect(sizeText(76229459)).toBe("72.7 MB");
    expect(sizeText(100)).toBe("1 KB");
    expect(whenText("not a time")).toBe("");
    expect(whenText("2026-11-14T03:05:00")).toBe("2026-11-14 03:05");
  });
});
