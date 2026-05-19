namespace BmesNgRateStandalone.Services;

/// <summary>Single definition of a menu item exposed in the nav menu.</summary>
public sealed record MenuItemDef(string Id, string Label, string Group);

/// <summary>
/// Central registry of navigable menus. The <c>Id</c> is the stable key used in the
/// MenuPermissions table; changing an Id requires a data migration.
/// </summary>
public static class AppMenus
{
    // IDs: keep stable, referenced by DB rows.
    public const string NgRate             = "ng-rate";
    public const string NgRateAll          = "ng-rate-all";
    public const string BmesWorkerStatus   = "bmes-worker-status";
    public const string BmesMakeModelGroup = "bmes-make-model-group";
    public const string BmesSetting        = "bmes-setting";
    public const string BmesRoutingTable   = "bmes-routing-table";
    public const string BmesReasonTable    = "bmes-reason-table";
    public const string BmesFCost          = "bmes-f-cost";
    public const string Schedule           = "schedule";
    public const string GraphMaker         = "graph-maker";
    public const string DiInput            = "di-input";
    public const string DiInputTest        = "di-input-test";
    public const string DiDb               = "di-db";
    public const string DiBatch            = "di-batch";
    public const string DiAnalysis         = "di-analysis";
    public const string DiModelAnalysis    = "di-model-analysis";
    public const string DiValidation       = "di-validation";
    public const string DiAsk              = "di-ask";
    public const string Report             = "report";
    public const string Translate          = "translate";
    public const string AdminUsers         = "admin-users";
    public const string AdminSettings      = "admin-settings";
    public const string AdminAiUsages      = "admin-ai-usages";
    public const string AdminPaths         = "admin-paths";
    public const string TestExcelConverter = "test-excel-converter";
    public const string AdminDbQuery       = "admin-db-query";

    public static readonly MenuItemDef[] All =
    [
        new(NgRate,             "Report NG RATE",           "BMES"),
        new(BmesWorkerStatus,   "Worker Status",            "BMES"),
        new(BmesMakeModelGroup, "Model Group",              "BMES"),
        new(BmesSetting,        "BMES Setting",             "BMES"),
        new(BmesRoutingTable,   "Routing Table",            "BMES"),
        new(BmesReasonTable,    "Reason Table",             "BMES"),
        new(BmesFCost,          "F-Cost",                   "BMES"),
        new(Schedule,           "GN LAB Schedule",          "Tools"),
        new(GraphMaker,         "Graph Maker",              "Tools"),
        new(DiInput,            "DI - Input Data",          "Data Inference"),
        new(DiInputTest,        "DI - Input Data (Test)",   "Data Inference"),
        new(DiDb,               "DI - DB Data",             "Data Inference"),
        new(DiBatch,            "DI - AI Batch",            "Data Inference"),
        new(DiAnalysis,         "DI - Aggregated Results",  "Data Inference"),
        new(DiModelAnalysis,    "DI - Each Model Analysis", "Data Inference"),
        new(DiValidation,       "DI - Validation",          "Data Inference"),
        new(DiAsk,              "DI - Ask AI",              "Data Inference"),
        new(Report,             "Report",                   "Data Inference"),
        new(Translate,          "Translate",                "Tools"),
        new(AdminUsers,         "Users",                    "Admin"),
        new(AdminSettings,      "Settings",                 "Admin"),
        new(AdminAiUsages,      "AI Usages",                "Admin"),
        new(AdminPaths,         "App Paths",                "Admin"),
        new(TestExcelConverter, "Test Excel Converter",     "Admin"),
        new(AdminDbQuery,       "DB Query",                 "Admin"),
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
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesSetting, BmesRoutingTable, BmesReasonTable, BmesFCost, Schedule,
            GraphMaker, DiInput, DiDb, DiBatch, DiAnalysis, DiModelAnalysis, DiValidation, DiAsk, Report, Translate
        ],
        [AppRoles.ManagerAi] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesSetting, BmesRoutingTable, BmesReasonTable, BmesFCost, Schedule,
            GraphMaker, DiInput, DiDb, DiBatch, DiAnalysis, DiModelAnalysis, DiValidation, DiAsk, Report, Translate
        ],
        [AppRoles.Leader] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesRoutingTable, BmesReasonTable, BmesFCost, Schedule,
            GraphMaker, DiInput, DiDb, DiBatch, DiAnalysis, DiModelAnalysis, DiValidation, DiAsk, Report, Translate
        ],
        [AppRoles.Editor] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesRoutingTable, BmesReasonTable, BmesFCost, Schedule,
            GraphMaker, DiInput, DiDb, DiBatch, DiAnalysis, DiModelAnalysis, DiValidation, DiAsk, Report, Translate
        ],
        [AppRoles.Viewer] =
        [
            NgRate, BmesWorkerStatus, BmesFCost, Schedule, GraphMaker, DiDb, DiAnalysis, DiModelAnalysis, DiValidation, DiAsk, Report, Translate
        ],
    };
}
