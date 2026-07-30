namespace JinoSupporter.Web.Services;

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
    public const string BmesTest3          = "bmes-test3";
    public const string BmesTest4          = "bmes-test4";
    public const string BmesTest5          = "bmes-test5";
    public const string BmesLpa            = "bmes-lpa";
    public const string QrBakoData         = "qr-bako-data";
    public const string BmesDailyReport    = "bmes-daily-report";
    // Legacy keys retained only for existing DB permission rows.
    public const string BmesTest           = "bmes-test";
    public const string BmesTest2          = "bmes-test2";
    public const string GraphMaker         = "graph-maker";
    public const string DiInput            = "di-input";
    // Legacy key retained only for existing DB permissions and migrations.
    public const string DiInputTest        = "di-input-test";
    public const string DiInputBatch       = "di-input-batch";
    public const string DiAiPrompt         = "di-ai-prompt";
    public const string DiDb               = "di-db";
    public const string DiAnalysis         = "di-analysis";
    public const string DiModelAnalysis    = "di-model-analysis";
    public const string DiCurrentProblem   = "di-current-problem";
    public const string DiValidation       = "di-validation";
    public const string DiAsk              = "di-ask";
    public const string DailyTestInput     = "daily-test-input";
    public const string Report             = "report";
    public const string Translate          = "translate";
    public const string PcDownload         = "pc-download";
    public const string AdminUsers         = "admin-users";
    public const string AdminSettings      = "admin-settings";
    public const string AdminAiUsages      = "admin-ai-usages";
    public const string AdminPaths         = "admin-paths";
    public const string TestExcelConverter = "test-excel-converter";
    public const string AdminDbQuery       = "admin-db-query";

    public static readonly MenuItemDef[] All =
    [
        new(NgRate,             "Report NG RATE",           "BMES"),
        new(BmesFCost,          "F-Cost",                   "BMES"),
        new(BmesWorkerStatus,   "Worker Status",            "BMES"),
        new(BmesMakeModelGroup, "Model Group",              "BMES"),
        new(BmesSetting,        "Setting",                  "Setting"),
        new(BmesRoutingTable,   "Routing Table",            "BMES"),
        new(BmesReasonTable,    "Reason Table",             "BMES"),
        new(BmesTest3,          "Test 3",                   "BMES"),
        new(BmesTest4,          "Test 4",                   "BMES"),
        new(BmesTest5,          "BOM & Drawing",            "BMES"),
        new(BmesLpa,            "LPA",                      "BMES"),
        new(QrBakoData,         "QR BAKO DATA",             "BMES"),
        new(BmesDailyReport,    "DAILY REPORT",             "BMES"),
        new(GraphMaker,         "Graph Maker",              "Tools"),
        new(DiInputBatch,       "INPUT DATA (BATCH)",       "Test Data Analysis"),
        new(DiDb,               "Result",                   "Test Data Analysis"),
        new(DiAsk,              "Ask AI",                   "Test Data Analysis"),
        new(DailyTestInput,     "Input Data",               "Daily Test Data"),
        new(Report,             "Report",                   "Test Data Analysis"),
        new(Translate,          "Translate",                "Tools"),
        new(PcDownload,         "PC Download",              "Tools"),
        new(AdminUsers,         "Users",                    "Admin"),
        new(AdminAiUsages,      "AI Usages",                "Admin"),
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
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesSetting, BmesRoutingTable, BmesReasonTable, BmesFCost, BmesTest3, BmesTest4, BmesTest5,
            QrBakoData, BmesDailyReport, GraphMaker, DiInputBatch, DiDb, DiAsk, DailyTestInput, Report, Translate, PcDownload
        ],
        [AppRoles.ManagerAi] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesSetting, BmesRoutingTable, BmesReasonTable, BmesFCost, BmesTest3, BmesTest4, BmesTest5,
            QrBakoData, BmesDailyReport, GraphMaker, DiInputBatch, DiDb, DiAsk, DailyTestInput, Report, Translate, PcDownload
        ],
        [AppRoles.Leader] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesRoutingTable, BmesReasonTable, BmesFCost, BmesTest3, BmesTest4, BmesTest5,
            QrBakoData, BmesDailyReport, GraphMaker, DiInputBatch, DiDb, DiAsk, DailyTestInput, Report, Translate, PcDownload
        ],
        [AppRoles.Editor] =
        [
            NgRate, BmesWorkerStatus, BmesMakeModelGroup, BmesRoutingTable, BmesReasonTable, BmesFCost, BmesTest3, BmesTest4, BmesTest5,
            QrBakoData, BmesDailyReport, GraphMaker, DiInputBatch, DiDb, DiAsk, DailyTestInput, Report, Translate, PcDownload
        ],
        [AppRoles.Viewer] =
        [
            NgRate, BmesWorkerStatus, BmesFCost, QrBakoData, BmesDailyReport, GraphMaker, DiDb, DiAsk, DailyTestInput, Report, Translate, PcDownload
        ],
    };
}
