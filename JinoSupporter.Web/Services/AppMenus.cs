namespace JinoSupporter.Web.Services;

/// <summary>Single definition of a menu item exposed in the nav menu.</summary>
public sealed record MenuItemDef(string Id, string Label, string Group);

/// <summary>
/// Central registry of navigable menus. The <c>Id</c> is the stable key used in the
/// MenuPermissions table; changing an Id requires a data migration.
/// </summary>
public static class AppMenus
{
    // IDs — keep stable, referenced by DB rows.
    public const string NgRate           = "ng-rate";
    public const string NgRateAll        = "ng-rate-all";
    public const string BmesWorkerStatus = "bmes-worker-status";
    public const string BmesMakeModel    = "bmes-make-model";
    public const string BmesSetting      = "bmes-setting";
    public const string Schedule         = "schedule";
    public const string DiInput          = "di-input";
    public const string DiDb             = "di-db";
    public const string DiBatch          = "di-batch";
    public const string DiAnalysis       = "di-analysis";
    public const string DiAsk            = "di-ask";
    public const string Report           = "report";
    public const string Translate        = "translate";
    public const string AdminUsers       = "admin-users";
    public const string AdminSettings    = "admin-settings";

    public static readonly MenuItemDef[] All =
    [
        new(NgRate,           "Report NG RATE",       "BMES"),
        new(NgRateAll,        "Report NG RATE All",   "BMES"),
        new(BmesWorkerStatus, "Worker Status",        "BMES"),
        new(BmesMakeModel,    "Make Model Data",  "BMES"),
        new(BmesSetting,      "BMES Setting",     "BMES"),
        new(Schedule,         "Schedule",         "Tools"),
        new(DiInput,          "DI — Input Data",  "Data Inference"),
        new(DiDb,             "DI — DB Data",     "Data Inference"),
        new(DiBatch,          "DI — AI Batch",    "Data Inference"),
        new(DiAnalysis,       "DI — Analysis",    "Data Inference"),
        new(DiAsk,            "DI — Ask AI",      "Data Inference"),
        new(Report,           "Report",           "Data Inference"),
        new(Translate,        "Translate",        "Tools"),
        new(AdminUsers,       "Users",            "Admin"),
        new(AdminSettings,    "Settings",         "Admin"),
    ];

    /// <summary>
    /// Sensible default permissions used when seeding a fresh DB. Admin always
    /// gets everything (enforced in <c>MenuPermissionService.IsAllowed</c>).
    /// </summary>
    public static readonly Dictionary<string, string[]> DefaultsByRole = new()
    {
        [AppRoles.Admin] = All.Select(m => m.Id).ToArray(),
        [AppRoles.Manager] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModel, BmesSetting, Schedule,
            DiInput, DiDb, DiBatch, DiAnalysis, DiAsk, Report, Translate
        ],
        [AppRoles.Leader] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModel, Schedule,
            DiInput, DiDb, DiBatch, DiAnalysis, DiAsk, Report, Translate
        ],
        [AppRoles.Editor] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModel, Schedule,
            DiInput, DiDb, DiBatch, DiAnalysis, DiAsk, Report, Translate
        ],
        [AppRoles.Viewer] =
        [
            NgRate, BmesWorkerStatus, Schedule, DiDb, DiAnalysis, DiAsk, Report, Translate
        ],
    };
}
