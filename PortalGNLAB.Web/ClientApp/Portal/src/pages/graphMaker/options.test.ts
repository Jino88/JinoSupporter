import { describe, it, expect } from "vitest";
import { inputMode, graphFields, graphDefaults } from "./options";
describe("graph request options",()=>{it("maps specialized graph modes",()=>{expect([0,1,4,5,6,7].map(inputMode)).toEqual([0,0,0,2,3,4]);});it("provides a distinct default for every control",()=>{expect(new Set(graphFields.map(f=>f.key)).size).toBe(graphFields.length);for(const f of graphFields)expect(Object.hasOwn(graphDefaults,f.key)).toBe(true);});});
