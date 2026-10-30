// Pure helpers of the Setting page (pages/Setting.tsx), kept apart so they can be tested.

/** The models the API key can use, plus the saved one, so a model that dropped out of the list
 *  stays selected instead of silently switching on the next Save. */
export function modelOptions(available: string[], saved: string): string[] {
  const has = available.some((m) => m.toLowerCase() === saved.toLowerCase());
  return has || !saved ? available : [saved, ...available];
}

/** Windows path join used for the folder preview (AppPathsService.GetNgRateDirectory & co.). */
export function joinPath(root: string, name: string): string {
  const r = root.trim();
  if (!r) return name;
  return /[\\/]$/.test(r) ? r + name : r + "\\" + name;
}

export interface WorkFolders {
  ngRate: string;
  fCost: string;
  workerStatus: string;
  serverTableTarget: string;
}

/** The folders every page uses under the working folder, for what is typed (blank = default). */
export function workFolders(
  typedRoot: string,
  defaults: { defaultDbRootDirectory: string; ngRateFolderName: string; fCostFolderName: string; workerStatusFolderName: string },
): WorkFolders {
  const root = typedRoot.trim() || defaults.defaultDbRootDirectory;
  const ngRate = joinPath(root, defaults.ngRateFolderName);
  return {
    ngRate,
    fCost: joinPath(root, defaults.fCostFolderName),
    workerStatus: joinPath(root, defaults.workerStatusFolderName),
    serverTableTarget: joinPath(ngRate, "ngrate_settings.db"),
  };
}

/** Same default as the Blazor page: January 1st two years ago. */
export function defaultHistoryFrom(today: Date = new Date()): string {
  return `${today.getFullYear() - 2}-01-01`;
}
