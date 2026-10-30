import { describe, expect, it } from "vitest";
import { defaultHistoryFrom, joinPath, modelOptions, workFolders } from "./settingLogic";

describe("modelOptions", () => {
  it("keeps the list when the saved model is in it (any case)", () => {
    expect(modelOptions(["gpt-a", "gpt-b"], "GPT-B")).toEqual(["gpt-a", "gpt-b"]);
  });
  it("puts a saved model that dropped out of the list first", () => {
    expect(modelOptions(["gpt-a"], "gpt-old")).toEqual(["gpt-old", "gpt-a"]);
  });
  it("offers the saved model while the list is empty", () => {
    expect(modelOptions([], "gpt-5.6-luna")).toEqual(["gpt-5.6-luna"]);
  });
});

describe("workFolders", () => {
  const defaults = {
    defaultDbRootDirectory: "C:\\Portal\\db",
    ngRateFolderName: "01. NG RATE",
    fCostFolderName: "02. FCOST",
    workerStatusFolderName: "03. WK STATUS",
  };
  it("joins every folder under the typed root", () => {
    expect(workFolders("  D:\\work\\ ", defaults)).toEqual({
      ngRate: "D:\\work\\01. NG RATE",
      fCost: "D:\\work\\02. FCOST",
      workerStatus: "D:\\work\\03. WK STATUS",
      serverTableTarget: "D:\\work\\01. NG RATE\\ngrate_settings.db",
    });
  });
  it("uses the default root when nothing is typed", () => {
    expect(workFolders("", defaults).ngRate).toBe("C:\\Portal\\db\\01. NG RATE");
  });
  it("joinPath does not double the separator", () => {
    expect(joinPath("D:/a/", "b")).toBe("D:/a/b");
    expect(joinPath("D:\\a", "b")).toBe("D:\\a\\b");
  });
});

describe("defaultHistoryFrom", () => {
  it("is January 1st two years back", () => {
    expect(defaultHistoryFrom(new Date(2026, 9, 30))).toBe("2024-01-01");
  });
});
