import { describe, expect, it } from "vitest";
import { builtinTask, DEFAULT_VOCAB, type VocabItem } from "./planVocab";

describe("builtinTask", () => {
  const own: VocabItem[] = [
    { code: "Make Sample", name: "샘플 제작", aliases: [] },
    { code: "Send Sample", name: "샘플 발송", aliases: [] },
    { code: "Materials", name: "자재", aliases: [] },
    { code: "S2", name: "Shipping", aliases: ["시료 발송"] },
    { code: "SETUP LINE", name: "SETUP LINE", aliases: [] },
    { code: "출장자", name: "출장자 출장", aliases: [] },
  ];

  it("is the code itself for a code of the default list", () => {
    expect(builtinTask("sample-make", DEFAULT_VOCAB.tasks)).toBe("sample-make");
    expect(builtinTask("material-in", own)).toBe("material-in");
  });

  it("knows a code of its own by its code, name or other names, without case, spaces or punctuation", () => {
    expect(["Make Sample", "Send Sample", "Materials", "S2"].map((c) => builtinTask(c, own))).toEqual(["sample-make", "sample-send", "material-in", "sample-send"]);
  });

  it("is none for a task that is no built-in task, or no code", () => {
    expect(builtinTask("SETUP LINE", own)).toBeUndefined();
    expect(builtinTask("출장자", own)).toBeUndefined();
    expect(builtinTask(undefined, own)).toBeUndefined();
  });
});
