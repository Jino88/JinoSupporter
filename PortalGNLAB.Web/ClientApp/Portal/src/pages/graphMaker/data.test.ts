import { describe,it,expect } from "vitest";
import { normalizeGraphData } from "./data";
describe("graph records input",()=>{it("preserves pasted tables",()=>expect(normalizeGraphData("A\tB\n1\t2")).toBe("A\tB\n1\t2"));it("unions columns and preserves missing, zero and multiline values",()=>expect(normalizeGraphData(JSON.stringify({records:[{a:0,b:"line\nbreak"},null,{b:2,c:3}]}))).toBe("a\tb\tc\n0\tline break\t\n\t2\t3"));it("rejects unsupported JSON",()=>expect(()=>normalizeGraphData("{}")).toThrow("records"));});
