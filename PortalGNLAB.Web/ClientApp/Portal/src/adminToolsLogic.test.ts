import { describe, expect, it } from "vitest";
import { percent, allowedExcel } from "./adminToolsLogic";
describe("admin tools input", () => {
 it("parses percentages without treating blanks as zero", () => { expect(percent(" 42% ")).toBe(42); expect(percent("")).toBeNull(); expect(percent("unknown")).toBeNull(); expect(percent("Infinity")).toBeNull(); });
 it("accepts the four original Excel extensions", () => { for (const name of ["a.xlsx","A.XLSM","b.xlsb","c.xls"]) expect(allowedExcel(name)).toBe(true); expect(allowedExcel("a.xlsx.exe")).toBe(false); });
});
