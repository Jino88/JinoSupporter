using System.Text.Json;
using PortalGNLAB.Web.Services;
using PortalGNLAB.Web.Services.BmesReports;
using PortalGNLAB.Web.Services.BmesReports.Contracts;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

if (args.Length != 1)
    throw new ArgumentException("Pass the absolute worktree root as the only argument.");

string workspace = Path.GetFullPath(args[0]);
string webProject = Path.Combine(workspace, "PortalGNLAB.Web");
var failures = new List<string>();
int checkCount = 0;

void Check(bool condition, string name)
{
    checkCount++;
    Console.WriteLine($"{(condition ? "PASS" : "FAIL")} {name}");
    if (!condition) failures.Add(name);
}

string validToken = "0123456789abcdef0123456789abcdef";
Check(BmesReportHtmlExportService.IsValidToken(validToken), "lowercase hex token accepted");
Check(!BmesReportHtmlExportService.IsValidToken("ABCDEF"), "uppercase token rejected");
Check(!BmesReportHtmlExportService.IsValidToken("../../report.json"), "path traversal token rejected");
Check(!BmesReportHtmlExportService.IsValidToken(new string('a', 65)), "overlength token rejected");

DateTimeOffset now = new(2026, 8, 18, 12, 0, 0, TimeSpan.Zero);
Check(BmesReportHtmlExportService.IsWithinCompleteReportTtl(now.AddMinutes(-15), now), "15-minute TTL boundary accepted");
Check(!BmesReportHtmlExportService.IsWithinCompleteReportTtl(now.AddMinutes(-15).AddTicks(-1), now), "expired TTL rejected");
Check(!BmesReportHtmlExportService.IsWithinCompleteReportTtl(now.AddTicks(1), now), "future metadata timestamp rejected");
Check(
    BmesReportHtmlExportService.IsCurrentReportContract(2, BmesReportContract.SchemaVersion, BmesReportContract.CalculationVersion),
    "cache v2/current schema accepted");
Check(
    !BmesReportHtmlExportService.IsCurrentReportContract(1, BmesReportContract.SchemaVersion, BmesReportContract.CalculationVersion),
    "legacy cache format excluded from React data");
Check(
    !BmesReportHtmlExportService.IsCurrentReportContract(2, "unsupported", BmesReportContract.CalculationVersion),
    "unsupported schema excluded from React data");

string webRoot = Path.Combine(webProject, "wwwroot");
var environment = new TestWebHostEnvironment(webProject, webRoot);
var monitor = new MutableOptionsMonitor<BmesReportViewerOptions>(new()
{
    ReactViewerEnabled = true,
    ViewerAssetVersion = "host-check-1",
});
var bootstrap = new BmesReportViewerBootstrap(
    environment,
    monitor,
    NullLogger<BmesReportViewerBootstrap>.Instance);
var artifacts = new BmesReportArtifacts(
    validToken,
    "report.html",
    "report.json",
    "cache.json",
    "cache-key",
    now,
    IsCurrentContract: true);

Check(bootstrap.Select(artifacts, forceLegacy: false).Mode == BmesReportViewerMode.React, "flag-on/current JSON/assets selects React");
Check(bootstrap.Select(artifacts, forceLegacy: true).Mode == BmesReportViewerMode.Legacy, "explicit legacy query selects fallback");
monitor.CurrentValue.ReactViewerEnabled = false;
Check(bootstrap.Select(artifacts, forceLegacy: false).Mode == BmesReportViewerMode.Legacy, "flag-off selects fallback");
monitor.CurrentValue.ReactViewerEnabled = true;
Check(bootstrap.Select(artifacts with { ReportJsonPath = null }, forceLegacy: false).Mode == BmesReportViewerMode.Legacy, "missing JSON selects fallback");
var missingAssetBootstrap = new BmesReportViewerBootstrap(
    new TestWebHostEnvironment(webProject, Path.Combine(webProject, "missing-wwwroot")),
    monitor,
    NullLogger<BmesReportViewerBootstrap>.Instance);
Check(missingAssetBootstrap.Select(artifacts, forceLegacy: false).Mode == BmesReportViewerMode.Legacy, "missing assets select fallback");

BmesReportBootstrapDocument document = bootstrap.BuildReactDocument(validToken);
Check(document.Html.Contains($"/report/bmes/data/{validToken}", StringComparison.Ordinal), "bootstrap uses same-origin data route");
Check(document.Html.Contains($"/report/bmes/view/{validToken}?legacy=true", StringComparison.Ordinal), "runtime failure has same-token legacy route");
Check(document.Html.Contains("/bmes-report/bmes-report.js?v=host-check-1", StringComparison.Ordinal), "ESM asset has cache version query");
Check(document.Html.Contains("/bmes-report/bmes-report.css?v=host-check-1", StringComparison.Ordinal), "CSS asset has cache version query");
Check(document.ContentSecurityPolicy.Contains("frame-ancestors 'self'", StringComparison.Ordinal), "bootstrap CSP preserves same-origin frame boundary");

string programSource = File.ReadAllText(Path.Combine(webProject, "Program.cs"));
int viewRoute = programSource.IndexOf("/report/bmes/view/{token}", StringComparison.Ordinal);
int dataRoute = programSource.IndexOf("/report/bmes/data/{token}", StringComparison.Ordinal);
int viewAccess = programSource.IndexOf("GetBmesReportAccessFailure", viewRoute, StringComparison.Ordinal);
int viewResolve = programSource.IndexOf("ResolveReportArtifacts", viewRoute, StringComparison.Ordinal);
int dataAccess = programSource.IndexOf("GetBmesReportAccessFailure", dataRoute, StringComparison.Ordinal);
int dataResolve = programSource.IndexOf("ResolveReportArtifacts", dataRoute, StringComparison.Ordinal);
Check(viewRoute >= 0 && viewAccess > viewRoute && viewAccess < viewResolve, "view route gates access before token resolution");
Check(dataRoute >= 0 && dataAccess > dataRoute && dataAccess < dataResolve, "data route gates access before token resolution");
Check(programSource.Contains("AppMenus.NgRate", StringComparison.Ordinal) && programSource.Contains("AppMenus.BmesFCost", StringComparison.Ordinal), "route permission mirrors BMES Report menu gate");

string exportSource = File.ReadAllText(Path.Combine(webProject, "Services", "BmesReportHtmlExportService.cs"));
int jsonSerialization = exportSource.IndexOf("BmesReportJson.SerializeToUtf8Bytes", StringComparison.Ordinal);
int tokenCreation = exportSource.IndexOf("Guid.NewGuid().ToString", jsonSerialization, StringComparison.Ordinal);
int publish = exportSource.IndexOf("PublishCompletedReport(cacheKey, token)", tokenCreation, StringComparison.Ordinal);
int cleanup = exportSource.IndexOf("CleanupOldTokens(token)", publish, StringComparison.Ordinal);
Check(jsonSerialization >= 0 && jsonSerialization < tokenCreation, "report JSON serialization fails before token folder creation");
Check(tokenCreation >= 0 && tokenCreation < publish && publish < cleanup, "HTML/JSON/cache publish completes before old-token cleanup");
Check(exportSource.Contains("SemaphoreSlim GenerationLock", StringComparison.Ordinal), "generation remains serialized by SemaphoreSlim");

using (JsonDocument settings = JsonDocument.Parse(File.ReadAllText(Path.Combine(webProject, "appsettings.json"))))
{
    JsonElement section = settings.RootElement.GetProperty(BmesReportViewerOptions.SectionName);
    Check(!section.GetProperty("ReactViewerEnabled").GetBoolean(), "operational default selects legacy HTML");
    Check(!string.IsNullOrWhiteSpace(section.GetProperty("ViewerAssetVersion").GetString()), "asset version configured");
}

using (JsonDocument fixture = JsonDocument.Parse(File.ReadAllText(
           Path.Combine(webProject, "ClientApp", "BmesReport", "test", "fixtures", "report-v1.json"))))
{
    JsonElement root = fixture.RootElement;
    Check(root.GetProperty("tabs").EnumerateObject().Count() == 8, "host fixture retains eight tabs");
    Check(root.GetProperty("viewerDefaults").GetProperty("minimumPpm").GetDouble() == 500d, "host fixture retains Minimum PPM 500");
}

Check(!File.Exists(Path.Combine(webRoot, "bmes-report", "report.json")), "report JSON is not exposed under wwwroot");
Check(new FileInfo(Path.Combine(webRoot, "bmes-report", "bmes-report.js")).Length > 0, "built ESM asset is non-empty");
Check(new FileInfo(Path.Combine(webRoot, "bmes-report", "bmes-report.css")).Length > 0, "built CSS asset is non-empty");

// Menus retired 2026-10-29: Plan / Memo / Board, Test 3, Test 4, LPA, Analysis, Daily Test, Graph, Download.
string[] retiredMenus =
[
    AppMenus.TeamMemo, AppMenus.BmesTest3, AppMenus.BmesTest4, AppMenus.BmesLpa, AppMenus.DiInputBatch,
    AppMenus.DiDb, AppMenus.DiAsk, AppMenus.DailyTestInput, AppMenus.GraphMaker, AppMenus.PcDownload,
];
Check(!AppMenus.All.Any(m => retiredMenus.Contains(m.Id)), "retired menus left out of AppMenus.All");
Check(!AppMenus.DefaultsByRole.Values.Any(ids => ids.Any(retiredMenus.Contains)), "retired menus left out of role defaults");
var adminMenu = PortalGNLAB.Web.Endpoints.PortalShellEndpoints.MenuFor(AppRoles.Admin, (_, _) => true);
Check(!adminMenu.Any(m => retiredMenus.Contains(m.Id)), "React menu lists no retired menu");
Check(adminMenu.Any(m => m.Id == AppMenus.Report && m.Group == "BMES" && m.React && m.Href == "/app/report/bmes"), "React menu keeps BMES Report (ported to /app)");
string navMenuSource = File.ReadAllText(Path.Combine(webProject, "Components", "Layout", "NavMenu.razor"));
string[] retiredLinks =
[
    "/app/memo", "bmes/test3", "bmes/test4", "bmes/lpa", "input-data-batch", "data-inference/",
    "daily-test-data/input", "/app/tools/graph-maker", "/app/tools/pc-download",
];
Check(!retiredLinks.Any(navMenuSource.Contains), "classic rail links no retired page");
string instrumentSource = File.ReadAllText(Path.Combine(webProject, "Components", "Layout", "InstrumentLayout.razor"));
Check(!retiredLinks.Select(l => l.TrimStart('/')).Any(l => instrumentSource.Contains($"\"{l}", StringComparison.Ordinal)), "new UI rail links no retired page");
string homeSource = File.ReadAllText(Path.Combine(webProject, "Components", "Pages", "Home.razor"));
Check(homeSource.Contains("url=/bmes/daily-report", StringComparison.Ordinal) && !homeSource.Contains("memo", StringComparison.OrdinalIgnoreCase), "start page opens Daily Report, not Plan");
string settingSource = File.ReadAllText(Path.Combine(webProject, "Components", "Pages", "BmesSettingPage.razor"));
Check(!settingSource.Contains("Team Memo", StringComparison.Ordinal) && !settingSource.Contains("Moved to its own BMES menu", StringComparison.Ordinal), "Setting has no Team Memo DB or moved-table cards");

// NG Rate redesign: process type badges are coloured by type, the tables keep their markup.
Check(NgRateUi.ProcessTypeClass("Sub Assy") == "pt-sub" && NgRateUi.ProcessTypeClass("MAIN ASSY") == "pt-main"
      && NgRateUi.ProcessTypeClass("Function") == "pt-func" && NgRateUi.ProcessTypeClass("Final Visual") == "pt-vis"
      && NgRateUi.ProcessTypeClass("OQC") == "pt-oqc" && NgRateUi.ProcessTypeClass("IPG") == "pt-other"
      && NgRateUi.ProcessTypeClass(null) == "pt-other", "process type badge classes");
string ngRateSource = File.ReadAllText(Path.Combine(webProject, "Components", "Pages", "NgRatePage.razor"));
Check(!ngRateSource.Contains("<span class=\"pt-badge\">", StringComparison.Ordinal)
      && ngRateSource.Contains("NgRateUi.ProcessTypeClass(", StringComparison.Ordinal), "NG Rate badges use the type colour");

// Review Report: "+ New" posts blank dates; they must bind (not HTTP 400) and real dates must survive.
var webJson = new JsonSerializerOptions(JsonSerializerDefaults.Web);
var blankInput = JsonSerializer.Deserialize<PortalGNLAB.Web.Endpoints.ToolPageEndpoints.ReportInput>(
    """{"report":{"id":"","createdAt":"","updatedAt":"","images":[],"revisions":[],"doc":null},"request":""}""", webJson);
Check(blankInput is not null && blankInput.Report.CreatedAt == default && blankInput.Report.UpdatedAt == default, "new review report with blank dates binds");
var datedInput = JsonSerializer.Deserialize<PortalGNLAB.Web.Endpoints.ToolPageEndpoints.ReportInput>(
    """{"report":{"createdAt":"2026-10-30T08:19:00","updatedAt":"2026-10-30T08:20:00"}}""", webJson);
Check(datedInput?.Report.CreatedAt == new DateTime(2026, 10, 30, 8, 19, 0)
      && JsonSerializer.Serialize(datedInput.Report, webJson).Contains("\"updatedAt\":\"2026-10-30T08:20:00\"", StringComparison.Ordinal),
    "saved review report dates round-trip");

// /app index.html names the current bundle, so the browser must not keep an old copy after a restart.
var shellContext = new Microsoft.AspNetCore.Http.DefaultHttpContext();
PortalGNLAB.Web.Endpoints.PortalShellEndpoints.ShellFileOptions.OnPrepareResponse(
    new Microsoft.AspNetCore.StaticFiles.StaticFileResponseContext(shellContext, new NotFoundFileInfo("index.html")));
Check(shellContext.Response.Headers.CacheControl.ToString().Contains("no-cache", StringComparison.Ordinal), "React index.html is served no-cache");

// AIRLEAK (and GAUSS once the column exists) are dated by ImportedAt; the tester clock is not trusted. BAKO stays on TestTime.
var resolveTime = typeof(QrTestDataService).GetMethod("ResolveTestTimeColumn",
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
string? TimeColumn(QrTestTableDef table, params string[] columns) => (string?)resolveTime.Invoke(null, [table, columns]);
Check(TimeColumn(QrTestTables.Airleak, "id", "TestTime", "ProductID", "ImportedAt") == "ImportedAt", "AIRLEAK dates by ImportedAt");
Check(TimeColumn(QrTestTables.Gauss, "OTIME", "ProductID") == "OTIME"
      && TimeColumn(QrTestTables.Gauss, "OTIME", "ProductID", "ImportedAt") == "ImportedAt", "GAUSS dates by ImportedAt when present, else OTIME");
Check(TimeColumn(QrTestTables.Bako, "Id", "TestTime", "ProductID", "ImportedAt") == "TestTime", "BAKO keeps TestTime");
string summarySql = (string)typeof(QrSummaryService).GetField("SummarySql",
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!.GetValue(null)!;
Check(summarySql.Contains("a.[ImportedAt] >= @From", StringComparison.Ordinal)
      && summarySql.Contains("a.[ImportedAt]                                  AS [TestTime]", StringComparison.Ordinal)
      && !summarySql.Contains("a.[TestTime]", StringComparison.Ordinal)
      && summarySql.Contains("b.[TestTime] >= @From", StringComparison.Ordinal), "QR SUMMARY dates AIRLEAK by ImportedAt, BAKO by TestTime");

// Review Report: always English, whatever language the notes are in.
string reportInstructions = (string)typeof(ReviewReportService).GetField("Instructions",
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!.GetValue(null)!;
string revisionInstructions = (string)typeof(ReviewReportService).GetField("RevisionInstructions",
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!.GetValue(null)!;
Check(ReviewReportService.Languages.SequenceEqual(["English"])
      && reportInstructions.Contains("Write EVERY text in English", StringComparison.Ordinal)
      && !reportInstructions.Contains("Same as input", StringComparison.Ordinal)
      && revisionInstructions.Contains("translated into", StringComparison.Ordinal), "review report is always written in English");

// Review Report preview: each result cell names the condition / metric it came from, so it can be edited in place.
var previewDoc = new ReviewReportDoc
{
    Conditions = [new() { Id = "C1", Label = "Film M-20", Model = "M-20", Table = "Film M-20", Role = "test" }],
    Metrics =
    [
        new() { Id = "input", Name = "Input", Unit = "pcs", Kind = "input", Values = [new() { ArmId = "C1", Value = "930" }] },
        new() { Id = "ng_spl", Name = "NG SPL", Unit = "pcs", Kind = "ng", Values = [new() { ArmId = "C1", Value = "4" }] },
        new() { Id = "ng_rb", Name = "NG RB", Unit = "pcs", Kind = "ng", Values = [new() { ArmId = "C1", Value = "51" }] },
    ],
};
var previewResult = typeof(PortalGNLAB.Web.Endpoints.ToolPageEndpoints).GetMethod("Preview",
    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!.Invoke(null, [new ReviewReport { Doc = previewDoc }])!;
object previewValue = previewResult.GetType().GetProperty("Value")!.GetValue(previewResult)!;
using var previewJson = JsonDocument.Parse(JsonSerializer.Serialize(previewValue, webJson));
JsonElement previewTable = previewJson.RootElement[0];
static string? Prop(JsonElement e, string name) => e.TryGetProperty(name, out JsonElement v) ? v.GetString() : null;
string?[] mainIds = previewTable.GetProperty("columns").EnumerateArray().Select(c => Prop(c, "mainId")).ToArray();
Check(Prop(previewTable, "tableKey") == "Film M-20"
      && Prop(previewTable.GetProperty("rows")[0], "armId") == "C1"
      && mainIds.SequenceEqual(["input", "", "ng_spl", "ng_rb", "", ""])
      && previewTable.GetProperty("columns").EnumerateArray().All(c => Prop(c, "rateId") == ""),
    "review report preview names the condition and metric behind each cell (worked-out OK / Total NG / Total NG rate have none)");

// Review Report follows the author's form (130. TIU L5S3-01): Input | OK | NG items with the rate (%) under
// each | Total NG | Total NG rate, and the Excel sheet keeps the form's header box, sections and fonts.
var formDoc = new ReviewReportDoc
{
    Title = "TIU L5S3-01 [R] REPORT TEST NEW VP CARRIER SUPPORT JIG",
    ReportDate = "2026-08-28",
    Dept = "ME",
    Purpose = "Old JIG happen NG VP cutting offset => Test New JIG can improve or not",
    Content = ["Test New JIG and check NG rate"],
    ResultTitle = "Result check New Jig",
    Conclusion = "Check Jig after improve result NG cutting offset reduce. Need check more Q'ty.",
    Conditions =
    [
        new() { Id = "C1", Label = "Old Jig", Role = "control", Date = "2026-08-28" },
        new() { Id = "C2", Label = "New Jig (jig 1-17)", Role = "test", Date = "2026-08-28" },
        new() { Id = "C3", Label = "New Jig (jig 18-25)", Role = "test", Date = "2026-09-03" },
    ],
    Metrics =
    [
        new() { Id = "input", Name = "Input", Unit = "pcs", Kind = "input", Values = [new() { ArmId = "C1", Value = "160" }, new() { ArmId = "C2", Value = "2720" }, new() { ArmId = "C3", Value = "640" }] },
        new() { Id = "ng_vp", Name = "NG VP cutting offset", Unit = "pcs", Kind = "ng", Values = [new() { ArmId = "C1", Value = "8" }, new() { ArmId = "C2", Value = "3" }, new() { ArmId = "C3", Value = "1" }] },
        new() { Id = "ng_cut", Name = "No cutting", Unit = "pcs", Kind = "ng", Values = [new() { ArmId = "C1", Value = "0" }, new() { ArmId = "C2", Value = "1" }, new() { ArmId = "C3", Value = "0" }] },
    ],
};
ReviewReportResultTable formTable = ReviewReportService.ResultTables(formDoc).Single();
string[] armIds = ["C1", "C2", "C3"];
Check(formTable.Columns.Select(c => c.Header).SequenceEqual(["Input", "OK", "NG VP cutting offset", "No cutting", "Total NG", "Total NG rate"])
      && armIds.Select(formTable.Columns[1].MainText).SequenceEqual(["152", "2716", "639"])
      && armIds.Select(formTable.Columns[2].Below).SequenceEqual(["5.0%", "0.1%", "0.2%"])
      && armIds.Select(formTable.Columns[3].Below).SequenceEqual(["0.0%", "0.0%", "0.0%"])
      && armIds.Select(formTable.Columns[4].MainText).SequenceEqual(["8", "4", "1"]) && !formTable.Columns[4].HasBelow
      && armIds.Select(formTable.Columns[5].MainText).SequenceEqual(["5.0%", "0.1%", "0.2%"])
      && formTable.Columns.Select(c => c.Computed).SequenceEqual([false, true, false, false, true, true]),
    "review report result table follows the form: Input | OK | NG with rate (%) under | Total NG | Total NG rate");

using (var formBook = new ClosedXML.Excel.XLWorkbook(new MemoryStream(ReviewReportService.ExportXlsx(new ReviewReport { Title = "x", Doc = formDoc }))))
{
    var sheet = formBook.Worksheet(1);
    string[] sectionsInB = sheet.Column(2).CellsUsed().Select(c => c.GetString()).ToArray();
    var head = sheet.Search("No").FirstOrDefault(c => c.Address.ColumnNumber == 3);
    int headRow = head?.Address.RowNumber ?? 1;
    int dataRow = headRow + 1;
    Check(head is not null && sheet.Name == "Test"
          && sheet.Cell("B2").GetString() == "TITLE" && sheet.Cell("D2").GetString() == formDoc.Title && sheet.Cell("D2").Style.Font.FontSize == 24
          && sheet.Cell("L2").GetString() == "Dept" && sheet.Cell("M2").GetString() == "ME"
          && sheet.Cell("L3").GetString() == "Date" && sheet.Cell("M3").Style.NumberFormat.Format == "d-mmm"
          && sheet.Cell("L4").GetString() == "Marker"
          && new[] { "I. Purpose.", "II. Content.", "III. Result", "1. Result check New Jig", "IV. Decision" }.All(sectionsInB.Contains)
          && sectionsInB.Contains("=> " + formDoc.Conclusion)
          && sheet.Cell("B5").Style.Font.FontName == "Times New Roman" && sheet.Cell("B5").Style.Font.FontSize == 16
          && sheet.Range(headRow, 3, headRow, 11).Cells().Select(c => c.GetString())
                .SequenceEqual(["No", "Date", "Type", "Input", "OK", "NG VP cutting offset", "No cutting", "Total NG", "Total NG rate"])
          && sheet.Cell(dataRow, 4).Style.NumberFormat.Format == "d-mmm" && sheet.Cell(dataRow, 4).IsMerged()
          && sheet.Cell(dataRow + 1, 8).GetDouble() == 0.05 && sheet.Cell(dataRow + 1, 8).Style.NumberFormat.Format == "0.0%"
          && sheet.Cell(dataRow, 11).Style.NumberFormat.Format == "0.0%" && sheet.Cell(dataRow, 11).IsMerged()
          && sheet.Cell(dataRow + 2, 5).Style.Fill.BackgroundColor == ClosedXML.Excel.XLColor.FromHtml("#FFFF00")
          && sheet.Cell(dataRow + 2, 3).Style.Fill.BackgroundColor != ClosedXML.Excel.XLColor.FromHtml("#FFFF00"),
        "review report Excel keeps the form's header box, sections, font and result table");
}

// Review Report: the Codex API model and effort come from Setting, Codex Luna at low until one is saved.
string settingDb = Path.Combine(Path.GetTempPath(), $"host-check-{Guid.NewGuid():N}.db");
try
{
    var settingRepo = new WebRepository(Microsoft.Extensions.Configuration.MemoryConfigurationBuilderExtensions.AddInMemoryCollection(
        new Microsoft.Extensions.Configuration.ConfigurationBuilder(),
        new Dictionary<string, string?> { ["Database:Path"] = settingDb }).Build());
    var writer = new ReviewReportService(null!, settingRepo, NullLogger<ReviewReportService>.Instance);
    string before = writer.ModelLabel;
    settingRepo.SetSetting(ReviewReportService.ModelSetting, " gpt-5.6-terra ");
    settingRepo.SetSetting(ReviewReportService.EffortSetting, "medium");
    string settingPage = File.ReadAllText(Path.Combine(webProject, "Components", "Pages", "BmesSettingPage.razor"));
    Check(before == $"{CodexApiService.LunaModel} · low"
          && writer.Model == "gpt-5.6-terra" && writer.Effort == "medium"
          && settingPage.Contains("Repo.SetSetting(ReviewReportService.ModelSetting", StringComparison.Ordinal)
          && settingPage.Contains("Repo.SetSetting(ReviewReportService.EffortSetting", StringComparison.Ordinal),
        "review report model and effort are chosen in Setting");
}
finally
{
    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
    foreach (string f in new[] { settingDb, settingDb + "-wal", settingDb + "-shm" })
        if (File.Exists(f)) File.Delete(f);
}

// Production Input was retired: its menu id stays only for old DB rows and is offered to no role.
Check(AppMenus.DefaultsByRole.Values.All(menus => !menus.Contains(AppMenus.ProductionInput))
      && !AppMenus.All.Any(m => m.Id == AppMenus.ProductionInput),
    "production-input menu is not offered to any role");

// Working: per-model info + entries, read-only model groups, menu entry and one-time grant.
string workDb = Path.Combine(Path.GetTempPath(), $"working-{Guid.NewGuid():N}.db");
string glossDb = Path.Combine(Path.GetTempPath(), $"glossary-{Guid.NewGuid():N}.db");
string workGrantDb = Path.Combine(Path.GetTempPath(), $"working-grant-{Guid.NewGuid():N}.db");
try
{
    JsonElement WJ(string json) => JsonDocument.Parse(json).RootElement.Clone();
    WorkingEntrySave WBody(string kind = "dev", string date = "2026-11-01", string title = "T", int rev = 0) =>
        new(kind, date, title, "body", "Open", "High", "", "kim", "2026-11-30", rev);
    bool WConflicts(Action act)
    {
        try { act(); return false; }
        catch (WorkingConflictException) { return true; }
    }
    bool WBad(Action act)
    {
        try { act(); return false; }
        catch (ArgumentException) { return true; }
    }

    var working = new WorkingService(workDb, glossDb);

    // Model info
    WorkingModel none = working.GetModel("X316B");
    Check(none.Revision == 0 && none.Info.ValueKind == JsonValueKind.Object && none.UpdatedAt == "", "working: unknown model has empty info and revision 0");
    WorkingModel m1 = working.SaveModel("X316B", new(WJ("""{"customer":"A","n":[1]}"""), 0), "alice");
    Check(m1.Revision == 1 && m1.UpdatedBy == "alice" && m1.Info.GetProperty("customer").GetString() == "A", "working: first save creates revision 1");
    Check(WConflicts(() => working.SaveModel("x316b", new(WJ("{}"), 0), "bob")), "working: stale model revision conflicts (case-insensitive name)");
    WorkingModel m2 = working.SaveModel("x316b", new(WJ("""{"customer":"B"}"""), 1), "bob");
    Check(m2.Revision == 2 && working.GetModel("X316B").Info.GetProperty("customer").GetString() == "B", "working: matching revision saves and bumps");
    Check(WConflicts(() => working.SaveModel("NEW1", new(WJ("{}"), 3), "bob")), "working: new model must start at revision 0");
    Check(WBad(() => working.SaveModel(new string('a', 81), new(WJ("{}"), 0), "x")) && WBad(() => working.GetModel("  ")), "working: model name length validated");

    // Entries
    WorkingEntry e1 = working.CreateEntry("X316B", WBody(date: "2026-11-01", title: "first"), "alice");
    WorkingEntry e2 = working.CreateEntry("X316B", WBody("issue", "2026-11-05", "second"), "alice");
    WorkingEntry e3 = working.CreateEntry("X316B", WBody("review", "2026-11-05", "third"), "alice");
    List<WorkingEntry> entries = working.ListEntries("x316b");
    Check(entries.Select(e => e.Title).SequenceEqual(["third", "second", "first"]) && entries.All(e => e.Revision == 1 && e.CreatedBy == "alice"),
        "working: entries sorted by date desc then created desc");
    Check(working.ListEntries("OTHER").Count == 0, "working: entries are per model");
    Check(working.UpdateEntry("X316B", "nope", WBody(), "u") is null, "working: updating a missing entry returns null (404)");
    WorkingEntry? e1b = working.UpdateEntry("X316B", e1.Id, WBody("issue", "2026-11-02", "renamed", 1), "carol");
    Check(e1b is { Revision: 2, Kind: "issue", Title: "renamed", UpdatedBy: "carol", CreatedBy: "alice" }, "working: entry update bumps revision and keeps creator");
    Check(WConflicts(() => working.UpdateEntry("X316B", e1.Id, WBody(rev: 1), "dave")), "working: stale entry revision conflicts");
    Check(working.DeleteEntry("X316B", e2.Id) && !working.DeleteEntry("X316B", e2.Id) && working.ListEntries("X316B").Count == 2, "working: delete removes once, second delete is false (404)");
    Check(WBad(() => working.CreateEntry("X316B", WBody(kind: "bogus"), "u")), "working: invalid kind rejected");
    Check(WBad(() => working.CreateEntry("X316B", WBody(date: "2026/11/01"), "u")), "working: invalid date rejected");
    Check(WBad(() => working.CreateEntry("X316B", WBody(title: " "), "u")) && WBad(() => working.CreateEntry("X316B", WBody(title: new string('t', 201)), "u")), "working: title required and capped at 200");
    Check(WBad(() => working.CreateEntry("X316B", WBody() with { Body = new string('b', 20001) }, "u")), "working: body capped at 20000");

    // Issue register: per model + kind numbers, issue fields, dated progress notes.
    WorkingEntrySave IBody(string title, int rev = 0) => new("issue", "2026-11-03", title, "", "Open", "", "", "kim", "2026-11-20", rev,
        "PVT", "Quality", "FUNCTION", "Material", "", "");
    WorkingEntry i1 = working.CreateEntry("SEQ", IBody("one"), "alice");
    WorkingEntry i2 = working.CreateEntry("SEQ", IBody("two"), "alice");
    WorkingEntry d1 = working.CreateEntry("SEQ", WBody("dev", "2026-11-03", "dev one"), "alice");
    Check(i1.Seq == 1 && i2.Seq == 2 && d1.Seq == 1 && working.CreateEntry("SEQ2", IBody("other model"), "a").Seq == 1,
        "working: Seq counts per model and per kind");
    working.DeleteEntry("SEQ", i2.Id);
    Check(working.CreateEntry("SEQ", IBody("three"), "alice").Seq == 3, "working: Seq is not reused after a delete");
    WorkingEntry? i1b = working.UpdateEntry("SEQ", i1.Id, IBody("one", 1) with { DoneDate = "2026-11-08", FinalResult = "Mesh changed", Status = "Closed" }, "bob");
    Check(i1b is { Seq: 1, Phase: "PVT", Category: "Quality", Process: "FUNCTION", Cause: "Material", DoneDate: "2026-11-08", FinalResult: "Mesh changed", Due: "2026-11-20" },
        "working: issue fields saved and Seq kept on update");
    Check(working.UpdateEntry("SEQ", d1.Id, WBody("issue", "2026-11-03", "now an issue", 1), "bob") is { Seq: 4 }, "working: kind change takes the next number of the new kind");
    Check(WBad(() => working.CreateEntry("SEQ", IBody("x") with { DoneDate = "11/08" }, "u")) && WBad(() => working.CreateEntry("SEQ", IBody("x") with { Due = "soon" }, "u"))
          && working.CreateEntry("SEQ", IBody("no dates") with { Due = "", DoneDate = "" }, "u").Due == "",
        "working: due and done date are empty or yyyy-MM-dd");
    Check(WBad(() => working.CreateEntry("SEQ", IBody("x") with { Category = new string('c', 61) }, "u")), "working: issue type capped at 60");

    WorkingEntryLog? n2 = working.CreateLog("SEQ", i1.Id, new("2026-11-06", " second day "), "alice");
    WorkingEntryLog? n1 = working.CreateLog("seq", i1.Id, new("2026-11-04", "first day"), "alice");
    Check(n1 is { Revision: 1, CreatedBy: "alice" } && n2 is { Text: "second day" }, "working: progress note created (text trimmed)");
    Check(working.CreateLog("SEQ2", i1.Id, new("2026-11-04", "x"), "a") is null && working.CreateLog("SEQ", "nope", new("2026-11-04", "x"), "a") is null,
        "working: a note needs an entry of the same model (404)");
    Check(WBad(() => working.CreateLog("SEQ", i1.Id, new("2026/11/04", "x"), "a")) && WBad(() => working.CreateLog("SEQ", i1.Id, new("2026-11-04", "  "), "a"))
          && WBad(() => working.CreateLog("SEQ", i1.Id, new("2026-11-04", new string('t', 5001)), "a")), "working: note date and text validated");
    WorkingEntry listed = working.ListEntries("SEQ").Single(e => e.Id == i1.Id);
    Check(listed.Logs.Select(l => l.Text).SequenceEqual(["first day", "second day"]) && listed.Revision == 2,
        "working: entries list carries notes oldest first; notes do not bump the entry revision");
    Check(working.ListEntries("SEQ").Where(e => e.Id != i1.Id).All(e => e.Logs.Count == 0), "working: notes stay on their own entry");
    WorkingEntryLog? n1b = working.UpdateLog("SEQ", i1.Id, n1!.Id, new("2026-11-07", "moved", 1), "bob");
    Check(n1b is { Revision: 2, Date: "2026-11-07", Text: "moved", UpdatedBy: "bob" }
          && working.ListEntries("SEQ").Single(e => e.Id == i1.Id).Logs.Select(l => l.Text).SequenceEqual(["second day", "moved"]),
        "working: note update bumps its revision and re-sorts by date");
    Check(WConflicts(() => working.UpdateLog("SEQ", i1.Id, n1.Id, new("2026-11-07", "stale", 1), "c")), "working: stale note revision conflicts (409)");
    Check(working.UpdateLog("SEQ", i1.Id, "nope", new("2026-11-07", "x", 1), "c") is null, "working: updating a missing note returns null (404)");
    Check(working.DeleteLog("SEQ", i1.Id, n2!.Id) && !working.DeleteLog("SEQ", i1.Id, n2.Id) && !working.DeleteLog("SEQ2", i1.Id, n1.Id),
        "working: note delete removes once and only within the model");
    Check(working.DeleteEntry("SEQ", i1.Id) && working.UpdateLog("SEQ", i1.Id, n1.Id, new("2026-11-07", "x", 2), "c") is null,
        "working: deleting an entry deletes its notes");

    // Note schedule / done date and pictures pasted into an issue or a note.
    WorkingEntry p1 = working.CreateEntry("PIC", IBody("pictures"), "alice");
    WorkingEntry p2 = working.CreateEntry("PIC", IBody("other issue"), "alice");
    WorkingEntryLog pn = working.CreateLog("PIC", p1.Id, new("2026-11-09", "measured", 0, " lee ", "2026-11-12", "2026-11-10"), "alice")!;
    Check(pn is { Owner: "lee", Due: "2026-11-12", DoneDate: "2026-11-10" } && pn.Images.Count == 0, "working: note owner, due and done date saved");
    Check(working.UpdateLog("PIC", p1.Id, pn.Id, new("2026-11-09", "measured", 1, "kim", "", ""), "bob") is { Owner: "kim", Due: "", DoneDate: "", Revision: 2 },
        "working: note schedule changed on update");
    Check(WBad(() => working.CreateLog("PIC", p1.Id, new("2026-11-09", "x", 0, "", "soon"), "a"))
          && WBad(() => working.CreateLog("PIC", p1.Id, new("2026-11-09", "x", 0, "", "", "11/10"), "a"))
          && WBad(() => working.CreateLog("PIC", p1.Id, new("2026-11-09", "x", 0, new string('o', 101)), "a")),
        "working: note due / done date are yyyy-MM-dd and owner is capped at 100");
    WorkingEntry detIssue = working.CreateEntry("DET", IBody("detail issue"), "alice");
    WorkingEntryLog detNote = working.CreateLog("DET", detIssue.Id, new("2026-11-09", " one line ", Detail: " line 2\nline 3 "), "alice")!;
    Check(detNote is { Text: "one line", Detail: "line 2\nline 3" }
          && working.UpdateLog("DET", detIssue.Id, detNote.Id, new("2026-11-09", "renamed", 1), "bob") is { Text: "renamed", Detail: "line 2\nline 3" }
          && working.UpdateLog("DET", detIssue.Id, detNote.Id, new("2026-11-09", "renamed", 2, Detail: ""), "bob") is { Detail: "" },
        "working: note detail saved trimmed; Detail null keeps it on update, \"\" clears it");
    Check(WBad(() => working.CreateLog("DET", detIssue.Id, new("2026-11-09", "two\nlines"), "a"))
          && WBad(() => working.CreateLog("DET", detIssue.Id, new("2026-11-09", "a\rb"), "a"))
          && WBad(() => working.CreateLog("DET", detIssue.Id, new("2026-11-09", new string('t', 501)), "a"))
          && WBad(() => working.CreateLog("DET", detIssue.Id, new("2026-11-09", "x", Detail: new string('d', 5001)), "a"))
          && working.CreateLog("DET", detIssue.Id, new("2026-11-09", new string('t', 500), Detail: new string('d', 5000)), "a") is not null,
        "working: a note's issue is one line of up to 500 characters; its detail up to 5000");
    byte[] png = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3];
    byte[] jpg = [0xFF, 0xD8, 0xFF, 0xE0, 9];
    List<WorkingImage>? onIssue = working.AddImages("pic", p1.Id, null, [("image.png", png), ("..\\x.jpg", jpg)], "alice");
    List<WorkingImage>? onNote = working.AddImages("PIC", p1.Id, pn.Id, [("note.png", png)], "bob");
    WorkingEntry withPics = working.ListEntries("PIC").Single(e => e.Id == p1.Id);
    Check(onIssue is [{ ContentType: "image/png", Size: 11, LogId: "" }, { FileName: "x.jpg", ContentType: "image/jpeg" }] && onNote is [{ CreatedBy: "bob" }]
          && withPics.Images.Select(i => i.Id).SequenceEqual(onIssue.Select(i => i.Id))
          && withPics.Logs.Single().Images.Single().Id == onNote[0].Id && onNote[0].LogId == pn.Id,
        "working: pictures kept in paste order under the issue and under the note");
    Check(working.ListEntries("PIC").Single(e => e.Id == p2.Id).Images.Count == 0, "working: pictures stay on their own issue");
    Check(working.GetImage(onIssue![0].Id) is { } got && got.ContentType == "image/png" && got.Data.SequenceEqual(png) && working.GetImage("nope") is null,
        "working: picture bytes round-trip");
    Check(WBad(() => working.AddImages("PIC", p1.Id, null, [("a.png", png), ("a.svg", "<svg/>"u8.ToArray())], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, [("fake.png", "MZ not a png"u8.ToArray())], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, [("big.png", [.. png, .. new byte[WorkingService.MaxImageBytes]])], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, [], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, Enumerable.Repeat(("a.png", png), 21).ToList(), "a"))
          && working.ListEntries("PIC").Single(e => e.Id == p1.Id).Images.Count == 2,
        "working: only PNG / JPEG / GIF / WebP / BMP bytes up to 10 MB, 1 to 20 at a time; a refused upload stores nothing");
    // Videos: told by their bytes, stored byte for byte, read back in ranges; HEIC / M4A (same box format) are refused.
    static byte[] Ftyp(string major, params string[] compat)
    {
        var box = new List<byte> { 0, 0, 0, (byte)(16 + 4 * compat.Length) };
        foreach (string s in new[] { "ftyp", major, "\0\0\0\0" }.Concat(compat)) box.AddRange(System.Text.Encoding.ASCII.GetBytes(s));
        return [.. box];
    }
    byte[] mp4 = [.. Ftyp("isom", "isom", "mp41"), .. Enumerable.Range(0, 300_000).Select(i => (byte)(i * 7))];
    byte[] mov = [.. Ftyp("qt  ", "qt  "), 5, 6];
    byte[] webm = [0x1A, 0x45, 0xDF, 0xA3, 0x9F, 0x42, 0x82, 0x84, .. "webm"u8, 1];
    byte[] mkv = [0x1A, 0x45, 0xDF, 0xA3, 0xA3, 0x42, 0x82, 0x88, .. "matroska"u8, 1];
    byte[] avi = [.. "RIFF"u8, 0, 0, 0, 0, .. "AVI LIST"u8];
    List<WorkingImage>? videos = working.AddImages("PIC", p1.Id, pn.Id, [("a.mp4", mp4), ("b.mov", mov), ("c.webm", webm), ("d.mkv", mkv), ("e.avi", avi)], "bob");
    Check(videos?.Select(v => v.ContentType).SequenceEqual(["video/mp4", "video/quicktime", "video/webm", "video/x-matroska", "video/x-msvideo"]) == true
          && videos[0].Size == mp4.Length && working.GetImage(videos[0].Id) is { } gotVideo && gotVideo.Data.SequenceEqual(mp4),
        "working: MP4 / MOV / WebM / MKV / AVI videos stored byte for byte");
    using (WorkingImageFile? open = working.OpenImage(videos![0].Id))
    {
        byte[] part = new byte[1000];
        open!.Data.Seek(200_000, SeekOrigin.Begin);
        open.Data.ReadExactly(part);
        Check(open.Data.CanSeek && open.Size == mp4.Length && part.SequenceEqual(mp4.AsSpan(200_000, 1000).ToArray()),
            "working: a stored video can be read from any point (byte ranges)");
    }
    Check(WBad(() => working.AddImages("PIC", p1.Id, null, [("p.heic", Ftyp("heic", "mif1", "heic"))], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, [("s.m4a", Ftyp("M4A ", "M4A ", "mp42", "isom"))], "a"))
          && WBad(() => working.AddImages("PIC", p1.Id, null, [("x.mkv", [0x1A, 0x45, 0xDF, 0xA3, 1, 2, 3])], "a"))
          && WBad(() => working.AddFiles("PIC", p1.Id, null, [new WorkingUpload("big.mp4", WorkingService.MaxVideoBytes + 1, () => new MemoryStream(mp4))], "a"))
          && WBad(() => working.AddFiles("PIC", p1.Id, null,
                 [new WorkingUpload("1.mp4", WorkingService.MaxVideoBytes, () => new MemoryStream(mp4)), new WorkingUpload("2.mp4", WorkingService.MaxVideoBytes, () => new MemoryStream(mp4)),
                  new WorkingUpload("3.mp4", 1, () => new MemoryStream(mp4))], "a"))
          && WBad(() => working.AddFiles("PIC", p1.Id, null, [new WorkingUpload("short.mp4", mp4.Length + 10, () => new MemoryStream(mp4))], "a"))
          && working.ListEntries("PIC").Single(e => e.Id == p1.Id).Images.Count == 2,
        "working: HEIC / M4A / unknown Matroska refused, videos up to 500 MB and 1 GB per upload, a short file refused; nothing stored");
    Check(working.AddImages("PIC", p1.Id, pn.Id, [("big.mp4", [.. Ftyp("mp42", "mp42"), .. new byte[WorkingService.MaxImageBytes]])], "a") is [{ ContentType: "video/mp4" }],
        "working: a video may be larger than the 10 MB picture limit");
    WorkingEntryLog otherNote = working.CreateLog("PIC", p2.Id, new("2026-11-09", "other"), "a")!;
    Check(working.AddImages("PIC", p1.Id, otherNote.Id, [("a.png", png)], "a") is null && working.AddImages("PIC", p1.Id, "nope", [("a.png", png)], "a") is null
          && working.AddImages("OTHER", p1.Id, null, [("a.png", png)], "a") is null,
        "working: pictures need an issue of the model and a note of that issue (404)");

    // Personal plan: model name ~personal:<login id>, only its owner
    Check(PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:alice", "alice")
          && PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk(PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalKey("alice"), "alice"),
        "working: personal key is open to its owner");
    Check(!PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:alice", "bob")
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:alice", "")
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:alice", null)
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:", "alice")
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~personal:alice2", "alice"),
        "working: personal key of someone else is refused");
    Check(PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~PERSONAL:Alice", "aLICE")
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("~Personal:Alice", "bob"),
        "working: personal key check ignores case");
    Check(PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("X316B", "bob")
          && PortalGNLAB.Web.Endpoints.WorkingEndpoints.PersonalOwnerOk("MSM-X526/X626B", "")
          && !PortalGNLAB.Web.Endpoints.WorkingEndpoints.IsPersonalModel("PIC") && PortalGNLAB.Web.Endpoints.WorkingEndpoints.IsPersonalModel("~personal:x"),
        "working: an ordinary model name is not a personal key");
    WorkingEntry pe = working.CreateEntry("~personal:alice", IBody("mine"), "alice");
    List<WorkingImage>? pImg = working.AddImages("~personal:alice", pe.Id, null, [("a.png", png)], "alice");
    Check(working.ImageModelName(pImg![0].Id) == "~personal:alice" && working.ImageModelName(onNote![0].Id) == "PIC" && working.ImageModelName("nope") is null,
        "working: a picture reports the model name of its entry");
    Check(WBad(() => working.AddTreeModel("~personal:bob", null, "a")) && WBad(() => working.AddTreeModel("~PERSONAL:x", null, "a")),
        "working: a personal key cannot be added to the model tree");
    Check(!working.DeleteImage("PIC", p2.Id, onIssue[0].Id) && !working.DeleteImage("OTHER", p1.Id, onIssue[0].Id)
          && working.DeleteImage("PIC", p1.Id, onIssue[0].Id) && !working.DeleteImage("PIC", p1.Id, onIssue[0].Id)
          && working.GetImage(onIssue[0].Id) is null && working.ListEntries("PIC").Single(e => e.Id == p1.Id).Images.Count == 1,
        "working: picture delete only under its own issue, once");
    Check(working.DeleteLog("PIC", p1.Id, pn.Id) && working.GetImage(onNote![0].Id) is null && working.GetImage(onIssue[1].Id) is not null,
        "working: deleting a note deletes its pictures only");
    WorkingEntryLog pn3 = working.CreateLog("PIC", p1.Id, new("2026-11-10", "again"), "a")!;
    string notePic = working.AddImages("PIC", p1.Id, pn3.Id, [("n.png", png)], "a")![0].Id;
    Check(working.DeleteEntry("PIC", p1.Id) && working.GetImage(onIssue[1].Id) is null && working.GetImage(notePic) is null,
        "working: deleting an issue deletes its pictures and its notes' pictures");

    // Setup Plan rows and the issues raised as their problems.
    WorkingSetupRowSave SRow(string process, int? no = null, string state = "", string status = "", int rev = 0) =>
        new(no, "", "SUB 1", process, "2026-11-20", state, "", "2026-11-27", "", status, rev);
    List<WorkingSetupRow> imported = working.ImportSetupRows("SET", [SRow("CD LOADING", 1, "x"), SRow("CD ETCHING", null, "△", "○"), SRow("CD BENDING", 5, "O", "O")], "alice");
    WorkingSetupRow s4 = working.CreateSetupRow("set", SRow("SP LOADING"), "bob");
    Check(imported.Select(r => (r.No, r.MachineState, r.SetupStatus)).SequenceEqual([(1, "X", ""), (2, "Δ", "O"), (5, "O", "O")]) && s4.No == 6
          && working.ListSetupRows("SET").Select(r => r.ProcessName).SequenceEqual(["CD LOADING", "CD ETCHING", "CD BENDING", "SP LOADING"])
          && working.ListSetupRows("OTHER").Count == 0,
        "working: setup rows numbered after the last, O / X / Δ look-alikes taken, listed by No. within the model");
    Check(WBad(() => working.ImportSetupRows("SET", [SRow("ok"), SRow("bad", state: "Y")], "a"))
          && WBad(() => working.CreateSetupRow("SET", SRow(" "), "a"))
          && WBad(() => working.CreateSetupRow("SET", SRow("x") with { SetupPlan = "27-Nov" }, "a"))
          && WBad(() => working.ImportSetupRows("SET", [], "a"))
          && working.ListSetupRows("SET").Count == 4,
        "working: setup state is O / X / Δ, process required, dates yyyy-MM-dd; a refused import adds nothing");
    WorkingSetupRow s1 = imported[0];
    Check(working.UpdateSetupRow("SET", s1.Id, SRow("CD LOADING", null, "O", "Δ", 1) with { SetupFinish = "2026-11-03" }, "c")
              is { No: 1, MachineState: "O", SetupStatus: "Δ", SetupFinish: "2026-11-03", Revision: 2, UpdatedBy: "c" }
          && WConflicts(() => working.UpdateSetupRow("SET", s1.Id, SRow("CD LOADING", rev: 1), "c"))
          && working.UpdateSetupRow("OTHER", s1.Id, SRow("x", rev: 2), "c") is null,
        "working: setup row update keeps its No., stale revision conflicts, other model 404");
    WorkingEntry prob = working.CreateEntry("SET", IBody("Not yet drawing laser") with { SetupRowId = s1.Id }, "alice");
    WorkingEntry plain = working.CreateEntry("SET", IBody("plain issue"), "alice");
    Check(prob.SetupRowId == s1.Id && plain.SetupRowId == "" && prob.Seq == 1 && plain.Seq == 2
          && working.UpdateEntry("SET", prob.Id, IBody("Not yet drawing laser", 1), "b") is { SetupRowId: var kept } && kept == s1.Id
          && working.CreateLog("SET", prob.Id, new("2026-11-09", "laser ordered"), "b") is not null,
        "working: an issue raised on a setup row keeps the row (same numbers, notes as any issue)");
    WorkingEntry otherRow = working.CreateEntry("OTHER", IBody("x"), "a");
    Check(WBad(() => working.CreateEntry("SET", IBody("x") with { SetupRowId = "nope" }, "a"))
          && WBad(() => working.CreateEntry("OTHER", IBody("x") with { SetupRowId = s1.Id }, "a"))
          && WBad(() => working.CreateEntry("SET", WBody("dev", "2026-11-03", "dev") with { SetupRowId = s1.Id }, "a"))
          && working.UpdateEntry("SET", plain.Id, IBody("plain issue", 1) with { SetupRowId = s4.Id }, "a") is { SetupRowId: var moved } && moved == s4.Id
          && otherRow.SetupRowId == "",
        "working: a problem needs a setup row of the same model and must be an issue");
    Check(working.DeleteSetupRow("SET", s1.Id) && !working.DeleteSetupRow("SET", s1.Id) && !working.DeleteSetupRow("OTHER", s4.Id)
          && working.ListEntries("SET").Single(e => e.Id == prob.Id) is { SetupRowId: "", Logs.Count: 1 }
          && working.ListEntries("SET").Single(e => e.Id == plain.Id).SetupRowId == s4.Id,
        "working: deleting a setup row keeps its problems as plain issues");

    // Sub-issues: one level under an issue.
    WorkingEntry par = working.CreateEntry("SET", IBody("parent issue") with { SetupRowId = s4.Id }, "alice");
    WorkingEntry sub = working.CreateEntry("SET", IBody("sub issue") with { ParentId = par.Id, SetupRowId = "ignored" }, "alice");
    WorkingEntry sub2 = working.CreateEntry("SET", IBody("sub issue 2") with { ParentId = par.Id }, "alice");
    WorkingEntry dev1 = working.CreateEntry("SET", WBody("dev", "2026-11-03", "dev item"), "alice");
    Check(sub.ParentId == par.Id && par.ParentId == "" && sub.SetupRowId == s4.Id && sub2.Seq == sub.Seq + 1 && sub.Kind == "issue"
          && working.ListEntries("SET").Single(e => e.Id == sub.Id).ParentId == par.Id,
        "working: a sub-issue is created under an issue, follows its setup row and keeps the model's numbering");
    Check(WBad(() => working.CreateEntry("SET", IBody("x") with { ParentId = otherRow.Id }, "a"))
          && WBad(() => working.CreateEntry("SET", IBody("x") with { ParentId = dev1.Id }, "a"))
          && WBad(() => working.CreateEntry("SET", IBody("x") with { ParentId = sub.Id }, "a"))
          && WBad(() => working.CreateEntry("SET", IBody("x") with { ParentId = "nope" }, "a"))
          && WBad(() => working.CreateEntry("SET", WBody("dev", "2026-11-03", "x") with { ParentId = par.Id }, "a")),
        "working: a parent must be a top-level issue of the same model (other model, non-issue, grandchild refused)");
    Check(WBad(() => working.UpdateEntry("SET", par.Id, IBody("parent issue", par.Revision) with { ParentId = par.Id }, "a"))
          && WBad(() => working.UpdateEntry("SET", par.Id, IBody("parent issue", par.Revision) with { ParentId = plain.Id }, "a"))
          && WBad(() => working.UpdateEntry("SET", plain.Id, IBody("plain issue", 2) with { ParentId = sub.Id }, "a")),
        "working: an issue cannot be its own parent, and an issue with sub-issues cannot move under another issue");
    Check(working.UpdateEntry("SET", sub.Id, IBody("sub issue renamed", sub.Revision), "a") is { ParentId: var keptParent } && keptParent == par.Id
          && working.UpdateEntry("SET", par.Id, IBody("parent issue", par.Revision) with { SetupRowId = "" }, "a") is { SetupRowId: "" }
          && working.ListEntries("SET").Where(e => e.ParentId == par.Id).All(e => e.SetupRowId == "" && e.Revision >= 2),
        "working: ParentId null keeps the parent; changing the parent's setup row moves its sub-issues too");
    WorkingEntry freed = working.CreateEntry("SET", IBody("to move"), "alice");
    Check(working.UpdateEntry("SET", freed.Id, IBody("to move", freed.Revision) with { ParentId = par.Id }, "a") is { ParentId: var moved2 } && moved2 == par.Id
          && working.UpdateEntry("SET", freed.Id, IBody("to move", freed.Revision + 1) with { ParentId = "" }, "a") is { ParentId: "" },
        "working: an issue can be moved under a parent and back to top level with ParentId \"\"");
    int subRev = working.ListEntries("SET").Single(e => e.Id == sub.Id).Revision;
    Check(working.DeleteEntry("SET", par.Id)
          && working.ListEntries("SET").Single(e => e.Id == sub.Id) is { ParentId: "" } afterDelete && afterDelete.Revision == subRev + 1
          && working.ListEntries("SET").Single(e => e.Id == sub2.Id).ParentId == "",
        "working: deleting a parent keeps its sub-issues as top-level issues");

    // A WorkingEntries table from before the issue columns gets them added and numbered on open.
    string oldDb = Path.Combine(Path.GetTempPath(), $"working-old-{Guid.NewGuid():N}.db");
    try
    {
        using (var oc = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={oldDb};Pooling=False"))
        {
            oc.Open();
            using var oCmd = oc.CreateCommand();
            oCmd.CommandText = """
                CREATE TABLE WorkingEntries (Id TEXT PRIMARY KEY, ModelName TEXT NOT NULL COLLATE NOCASE, Kind TEXT NOT NULL, Date TEXT NOT NULL,
                  Title TEXT NOT NULL, Body TEXT NOT NULL DEFAULT '', Status TEXT NOT NULL DEFAULT '', Severity TEXT NOT NULL DEFAULT '',
                  Result TEXT NOT NULL DEFAULT '', Owner TEXT NOT NULL DEFAULT '', Due TEXT NOT NULL DEFAULT '', Revision INTEGER NOT NULL,
                  CreatedBy TEXT NOT NULL DEFAULT '', CreatedAt TEXT NOT NULL, UpdatedBy TEXT NOT NULL DEFAULT '', UpdatedAt TEXT NOT NULL);
                INSERT INTO WorkingEntries VALUES ('b','M','issue','2026-11-09','later','','Open','Major','','','2026-11-20',1,'','2026-11-09T02:00:00Z','','2026-11-09T02:00:00Z');
                INSERT INTO WorkingEntries VALUES ('a','M','issue','2026-11-09','earlier','','Open','Major','','','',1,'','2026-11-09T01:00:00Z','','2026-11-09T01:00:00Z');
                CREATE TABLE WorkingEntryLogs (Id TEXT PRIMARY KEY, EntryId TEXT NOT NULL, Date TEXT NOT NULL, Text TEXT NOT NULL, Revision INTEGER NOT NULL,
                  CreatedBy TEXT NOT NULL DEFAULT '', CreatedAt TEXT NOT NULL, UpdatedBy TEXT NOT NULL DEFAULT '', UpdatedAt TEXT NOT NULL);
                INSERT INTO WorkingEntryLogs VALUES ('n','b','2026-11-09','old note',1,'u','2026-11-09T03:00:00Z','u','2026-11-09T03:00:00Z');
                INSERT INTO WorkingEntryLogs VALUES ('m','a','2026-11-09',' first line ' || char(13) || char(10) || ' second' || char(10) || 'third ',1,'u','2026-11-09T03:00:00Z','u','2026-11-09T03:00:00Z');
                """;
            oCmd.ExecuteNonQuery();
        }
        var migrated = new WorkingService(oldDb, glossDb);
        List<WorkingEntry> old = migrated.ListEntries("M");
        Check(old.Single(e => e.Id == "a").Seq == 1 && old.Single(e => e.Id == "b").Seq == 2 && old.All(e => e.Phase == "" && e.DoneDate == "" && e.Images.Count == 0 && e.SetupRowId == "" && e.ParentId == "")
              && migrated.CreateEntry("M", IBody("new"), "u").Seq == 3 && migrated.CreateLog("M", "a", new("2026-11-10", "note"), "u") is not null,
            "working: old table gets the issue columns, existing rows numbered by creation");
        Check(old.Single(e => e.Id == "b").Logs is [{ Text: "old note", Owner: "", Due: "", DoneDate: "" }]
              && migrated.UpdateLog("M", "b", "n", new("2026-11-09", "old note", 1, "lee", "2026-11-11"), "u") is { Owner: "lee", Due: "2026-11-11" }
              && migrated.AddImages("M", "b", "n", [("a.png", png)], "u") is [_],
            "working: old notes table gets owner / due / done date; pictures table created");
        Check(old.Single(e => e.Id == "a").Logs is [{ Text: "first line", Detail: "second\nthird" }]
              && new WorkingService(oldDb, glossDb).ListEntries("M").Single(e => e.Id == "b").Logs is [{ Text: "old note", Detail: "" }],
            "working: a several-line old note keeps its first line as the issue and moves the rest to Detail");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { File.Delete(oldDb); } catch (IOException) { }
    }

    // Problems of a Setup Plan row saved with the row's LINE as Process get its process name, once.
    string[] procDbs = [.. Enumerable.Range(0, 3).Select(i => Path.Combine(Path.GetTempPath(), $"working-proc{i}-{Guid.NewGuid():N}.db"))];
    try
    {
        var before = new WorkingService(procDbs[0], glossDb);
        List<WorkingSetupRow> sr = before.ImportSetupRows("PROC", [
            new(1, "", "SUB 1", " CD ETCHING ", "", "", "", "", "", ""), new(2, "", "SUB 1", new string('P', 70), "", "", "", "", "", "")], "a");
        WorkingEntrySave SBody(string title, string process, string setupRowId) =>
            new("issue", "2026-11-09", title, "", "Open", "", "", "", "", 0, "MP", "SETUP", process, "", "", "", setupRowId);
        before.CreateEntry("PROC", SBody("from line", "SUB 1", sr[0].Id), "a");
        before.CreateEntry("PROC", SBody("long name", "SUB 1", sr[1].Id), "a");
        before.CreateEntry("PROC", SBody("own process", "WELDING", sr[0].Id), "a");
        before.CreateEntry("PROC", SBody("no row", "SUB 1", ""), "a");
        void CopyDb(string from, string to, bool clearMark)
        {
            Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
            File.Copy(from, to);
            if (!clearMark) return;
            using var pc = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={to};Pooling=False");
            pc.Open();
            using var pCmd = pc.CreateCommand();
            pCmd.CommandText = "DROP TABLE IF EXISTS WorkingMigrations";
            pCmd.ExecuteNonQuery();
        }
        CopyDb(procDbs[0], procDbs[1], clearMark: true); // as a database from before the fix
        var fixedSvc = new WorkingService(procDbs[1], glossDb);
        Dictionary<string, string> procOf = fixedSvc.ListEntries("PROC").ToDictionary(e => e.Title, e => e.Process);
        Check(procOf["from line"] == "CD ETCHING" && procOf["long name"] == new string('P', 60)
              && procOf["own process"] == "WELDING" && procOf["no row"] == "SUB 1",
            "working: setup problems with the LINE as Process get the row's process name");
        WorkingEntry relined = fixedSvc.ListEntries("PROC").Single(e => e.Title == "from line");
        fixedSvc.UpdateEntry("PROC", relined.Id, SBody("from line", "SUB 1", sr[0].Id) with { Revision = relined.Revision }, "a");
        CopyDb(procDbs[1], procDbs[2], clearMark: false);
        Check(new WorkingService(procDbs[2], glossDb).ListEntries("PROC").Single(e => e.Title == "from line").Process == "SUB 1",
            "working: the Process fix runs once, a line set by hand later is kept");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        foreach (string f in procDbs)
            foreach (string s in new[] { f, f + "-wal", f + "-shm" })
                try { File.Delete(s); } catch (IOException) { }
    }

    // Model groups: missing file reports an error instead of throwing.
    WorkingGroups missing = working.ListGroups();
    Check(missing.Groups.Count == 0 && !string.IsNullOrEmpty(missing.Error) && missing.Source == glossDb, "working: missing glossary file gives empty groups plus error");
    using (var g = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={glossDb}"))
    {
        g.Open();
        using var c = g.CreateCommand();
        c.CommandText = """
            CREATE TABLE model_groups (id TEXT, name TEXT, parent_id TEXT, models TEXT, sort INT, position INT, updated_at_utc TEXT);
            INSERT INTO model_groups VALUES ('g2','Beta','g1','[]',1,1,'t'), ('g1','Alpha',NULL,'["X316B","X400"]',0,0,'t');
            """;
        c.ExecuteNonQuery();
    }
    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
    WorkingGroups groups = working.ListGroups();
    Check(groups.Error is null && groups.Groups.Count == 2 && groups.Groups[0].Id == "g1" && groups.Groups[0].ParentId is null
          && groups.Groups[0].Models.SequenceEqual(["X316B", "X400"]) && groups.Groups[1].ParentId == "g1",
        "working: model_groups read from a temp sqlite");
    var badTable = new WorkingService(workDb, workDb).ListGroups();
    Check(badTable.Groups.Count == 0 && badTable.Error is not null, "working: glossary without model_groups gives error, not exception");

    // Tree settings: hidden list, added models (glossary DB read only)
    string treeDb = Path.Combine(Path.GetTempPath(), $"working-tree-{Guid.NewGuid():N}.db");
    try
    {
        var tree = new WorkingService(treeDb, glossDb);
        Check(tree.GetTree() is { Hidden: [], Added: [] }, "working tree: new DB has nothing hidden or added");
        WorkingTree h1 = tree.SetHidden([" X400 ", "x400", "", null, "Other", "  "], "alice");
        Check(h1.Hidden.Count == 2 && h1.Hidden.Contains("X400") && h1.Hidden.Contains("Other") && tree.GetTree().Hidden.Count == 2,
            "working tree: hidden list is trimmed, blanks dropped, duplicates merged");
        WorkingTree h2 = tree.SetHidden(["X316B"], "bob");
        Check(h2.Hidden is ["X316B"] && tree.SetHidden([], "bob").Hidden.Count == 0, "working tree: hidden list is replaced as a whole");
        Check(WBad(() => tree.SetHidden([new string('a', 101)], "a")) && WBad(() => tree.SetHidden(Enumerable.Range(0, 5001).Select(i => "m" + i), "a"))
              && tree.SetHidden(Enumerable.Range(0, 5000).Select(i => "m" + i), "a").Hidden.Count == 5000,
            "working tree: hidden names over 100 characters or more than 5000 names are rejected");
        tree.SetHidden([], "a");

        WorkingAddedModel a1 = tree.AddTreeModel(" NEW-1 ", "", "alice");
        WorkingAddedModel a2 = tree.AddTreeModel("NEW-2", "g1", "bob");
        Check(a1.Name == "NEW-1" && a1.GroupId == "" && a1.CreatedBy == "alice" && a1.CreatedAt != ""
              && a2.GroupId == "g1" && tree.GetTree().Added.Select(a => a.Name).SequenceEqual(["NEW-1", "NEW-2"]),
            "working tree: added models are saved in creation order");
        Check(WConflicts(() => tree.AddTreeModel("new-1", "", "a")) && WConflicts(() => tree.AddTreeModel("x316b", "", "a"))
              && WConflicts(() => tree.AddTreeModel(" X400", "g2", "a")),
            "working tree: an added name or a glossary model name gives 409");
        Check(WBad(() => tree.AddTreeModel("NEW-3", "nope", "a")) && WBad(() => tree.AddTreeModel("  ", "", "a"))
              && WBad(() => tree.AddTreeModel(new string('a', 101), "", "a")) && tree.GetTree().Added.Count == 2,
            "working tree: unknown group, blank or over-long name gives 400");
        tree.SetHidden(["new-2", "Gone"], "alice");
        Check(tree.GetTree().Hidden.Count == 2 && tree.SetHidden(["Gone"], "alice").Hidden is ["Gone"] && tree.GetTree().Added.Count == 2,
            "working tree: an added model can be hidden and keeps existing when unhidden");
        WorkingAddedModel up = tree.AddTreeModel("Gone", "", "carol");
        Check(up.Name == "Gone" && tree.GetTree() is { Hidden: ["Gone"] } t3 && t3.Added.Count == 3, "working tree: a hidden-only name can be added later");
        Check(tree.DeleteTreeModel("new-1") && !tree.DeleteTreeModel("NEW-1") && !tree.DeleteTreeModel("X316B") && !tree.DeleteTreeModel("nothing")
              && tree.GetTree().Added.Select(a => a.Name).SequenceEqual(["NEW-2", "Gone"]),
            "working tree: added model deleted, missing and glossary names are not found");
        Check(new WorkingService(treeDb, Path.Combine(Path.GetTempPath(), "no-glossary-" + Guid.NewGuid().ToString("N") + ".db")) is { } noGloss
              && WBad(() => noGloss.AddTreeModel("NEW-9", "g1", "a")) && noGloss.AddTreeModel("NEW-9", "", "a").Name == "NEW-9",
            "working tree: without a readable glossary only the empty group is accepted");

        // Sub-groups made in Manage and models moved into them
        WorkingUserGroup ug1 = tree.AddTreeGroup(" X5 series ", "g1", "alice");
        WorkingUserGroup ug2 = tree.AddTreeGroup("Deep", ug1.Id, "alice");
        WorkingUserGroup ug3 = tree.AddTreeGroup("Next", "g1", "bob");
        Check(ug1.Id.StartsWith("ug-") && ug1.Name == "X5 series" && ug1.ParentId == "g1" && ug1.Sort == 0 && ug3.Sort == 1 && ug2.ParentId == ug1.Id
              && tree.GetTree().Groups.Count == 3,
            "working tree: sub-groups are made under a glossary group or another sub-group, after their siblings");
        Check(WConflicts(() => tree.AddTreeGroup("x5 SERIES", "g1", "a")) && WBad(() => tree.AddTreeGroup("Z", "nope", "a"))
              && WBad(() => tree.AddTreeGroup("  ", "g1", "a")) && WBad(() => tree.AddTreeGroup(new string('a', 61), "g1", "a"))
              && tree.AddTreeGroup("X5 series", "g2", "a") is { ParentId: "g2" },
            "working tree: a sub-group name used under the same group gives 409; unknown parent, blank or long name 400");
        Check(tree.RenameTreeGroup(ug3.Id, "Next gen", "carol") is { Name: "Next gen" } && tree.RenameTreeGroup("ug-none", "x", "a") is null
              && WConflicts(() => tree.RenameTreeGroup(ug3.Id, "X5 SERIES", "a")),
            "working tree: a sub-group is renamed; missing gives null, a sibling's name 409");

        WorkingTree? mv1 = tree.MoveTreeModel("x316b", ug1.Id, "alice");
        WorkingTree? mv2 = tree.MoveTreeModel("NEW-2", ug2.Id, "alice");
        Check(mv1 is not null && mv2 is not null && mv2.Placed.Count == 1 && mv2.Placed[0].Name == "X316B" && mv2.Placed[0].GroupId == ug1.Id
              && mv2.Added.First(a => a.Name == "NEW-2").GroupId == ug2.Id,
            "working tree: a glossary model (kept in its glossary spelling) and an added model are moved into sub-groups");
        Check(tree.MoveTreeModel("nobody", ug1.Id, "a") is null && WBad(() => tree.MoveTreeModel("X400", "ug-none", "a"))
              && tree.MoveTreeModel("X400", "g2", "a") is { Placed.Count: 2 },
            "working tree: an unknown model is not found, an unknown group gives 400, a glossary group is allowed");
        tree.SetHidden(["X316B"], "a");
        tree.SetHidden([], "a");
        Check(tree.GetTree().Placed.Any(p => p.Name == "X316B"), "working tree: unhiding a moved model keeps it where it was moved");
        Check(tree.MoveTreeModel("X400", "", "a") is { Placed.Count: 1 } && tree.MoveTreeModel("NEW-2", "", "a")!.Added.First(a => a.Name == "NEW-2").GroupId == "",
            "working tree: moving to \"\" puts a model back where it came from");

        tree.MoveTreeModel("NEW-2", ug2.Id, "a");
        tree.AddTreeModel("NEW-5", ug1.Id, "a");
        Check(tree.DeleteTreeGroup(ug1.Id) && !tree.DeleteTreeGroup(ug1.Id) && tree.GetTree() is { } afterDel
              && afterDel.Placed.Count == 0 && afterDel.Added.First(a => a.Name == "NEW-5").GroupId == "g1"
              && afterDel.Groups.First(g => g.Id == ug2.Id).ParentId == "g1" && afterDel.Added.First(a => a.Name == "NEW-2").GroupId == ug2.Id
              && afterDel.Groups.All(g => g.Id != ug1.Id),
            "working tree: deleting a sub-group moves its glossary models back, its added models and sub-groups up");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { File.Delete(treeDb); } catch (IOException) { }
    }

    // A DB from before the tree table gets it on open.
    string preTreeDb = Path.Combine(Path.GetTempPath(), $"working-pretree-{Guid.NewGuid():N}.db");
    try
    {
        using (var pc = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={preTreeDb};Pooling=False"))
        {
            pc.Open();
            using var pCmd = pc.CreateCommand();
            pCmd.CommandText = "CREATE TABLE WorkingModels (ModelName TEXT PRIMARY KEY COLLATE NOCASE, InfoJson TEXT NOT NULL DEFAULT '{}', Revision INTEGER NOT NULL, UpdatedAt TEXT NOT NULL, UpdatedBy TEXT NOT NULL DEFAULT '');";
            pCmd.ExecuteNonQuery();
        }
        var pre = new WorkingService(preTreeDb, glossDb);
        Check(pre.GetTree() is { Hidden: [], Added: [] } && pre.AddTreeModel("OLD-1", "", "u").Name == "OLD-1",
            "working tree: an old DB gets the tree table on open");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { File.Delete(preTreeDb); } catch (IOException) { }
    }

    // BMES model check: candidates from catalog / groups, merged with bmes first
    WorkingModelCandidate Cand(string n, string src, string code = "") => new(n, code, src);
    var checkCatalog = WorkingBmesService.CatalogCandidates("tiu-c11",
        [new() { ProductName = "ASSY TIU-C11-18", ProductCode = "P-S-1" }, new() { ProductName = "OTHER", ProductCode = "P-S-2" }, new() { ProductName = "tiu-c11-20", ProductCode = "P-S-3" }]);
    Check(checkCatalog.Count == 2 && checkCatalog.All(c => c.Source == "catalog") && checkCatalog[0].ProductCode == "P-S-1",
        "working check: catalog candidates match by contained text, any case");
    var manyCatalog = Enumerable.Range(0, 15).Select(i => new BmesBomModelCandidate { ProductName = "ABC" + i, ProductCode = "c" + i }).Append(new() { ProductName = "ABC", ProductCode = "x" }).ToList();
    var capped = WorkingBmesService.CatalogCandidates("abc", manyCatalog);
    Check(capped.Count == 10 && capped[0].ProductName == "ABC", "working check: catalog candidates are capped at 10 with the exact name first");
    var checkGroups = WorkingBmesService.GroupCandidates("tiu-c11-18",
    [
        new() { Name = "UNIT", ProductGroup = "UNIT", MidGroups = [new() { Material = "TIU-C11-18", SubGroups = [new() { Name = "", LineShifts = ["TIU-C11-18-L_E21A", "TIU-C11-18-L_E22B", "PLAIN_X"] }] }] },
    ]);
    Check(checkGroups.Count >= 1 && checkGroups.All(c => c.Source == "group" && c.ProductName.Contains("TIU-C11-18", StringComparison.OrdinalIgnoreCase))
          && checkGroups.Select(c => c.ProductName.ToLowerInvariant()).Distinct().Count() == checkGroups.Count,
        "working check: group candidates come from materials and product names without duplicates");
    var merged = WorkingBmesService.MergeCandidates(" abc-1 ",
        [Cand("ABC-1", "bmes", "B1")], [Cand("abc-1", "catalog", "C1"), Cand("ABC-2", "catalog", "C2")], [Cand("Abc-2", "group"), Cand("ABC-3", "group")]);
    Check(merged.Exact && merged.Candidates.Select(c => c.Source + ":" + c.ProductName).SequenceEqual(["bmes:ABC-1", "catalog:ABC-2", "group:ABC-3"]),
        "working check: duplicates keep the bmes > catalog > group source, exact ignores case and outer spaces");
    var noExact = WorkingBmesService.MergeCandidates("ABC", [], [Cand("ABC-1", "catalog")], [Cand("ABC-2", "group")]);
    Check(!noExact.Exact && noExact.Candidates.Count == 2 && !WorkingBmesService.MergeCandidates("ABC", [], [], []).Exact,
        "working check: a longer name is not exact, nothing found is not exact");

    // Menu + one-time grant
    var wManager = PortalGNLAB.Web.Endpoints.PortalShellEndpoints.MenuFor(AppRoles.Manager, (role, id) => AppMenus.DefaultsByRole[role].Contains(id));
    var wViewer = PortalGNLAB.Web.Endpoints.PortalShellEndpoints.MenuFor(AppRoles.Viewer, (role, id) => AppMenus.DefaultsByRole[role].Contains(id));
    Check(wManager.Any(m => m.Id == AppMenus.Working && m.React && m.Href == "/app/working" && m.Label == "Working")
          && !wViewer.Any(m => m.Id == AppMenus.Working),
        "working menu is offered to Manager with the React href and not to Viewer");
    var wRepo = new WebRepository(Microsoft.Extensions.Configuration.MemoryConfigurationBuilderExtensions.AddInMemoryCollection(
        new Microsoft.Extensions.Configuration.ConfigurationBuilder(),
        new Dictionary<string, string?> { ["Database:Path"] = workGrantDb }).Build());
    wRepo.SetMenuPermission(AppRoles.Leader, AppMenus.NgRate, true);
    var wPerms = new MenuPermissionService(wRepo);
    bool wGranted = wPerms.IsAllowed(AppRoles.Leader, AppMenus.Working);
    wRepo.SetMenuPermission(AppRoles.Leader, AppMenus.Working, false);
    _ = new MenuPermissionService(wRepo);
    Check(wGranted && !wPerms.IsAllowed(AppRoles.Leader, AppMenus.Working) && !wRepo.GetMenuPermissionsForRole(AppRoles.Viewer).Contains(AppMenus.Working),
        "working is granted once to seeded roles and an admin removal sticks");

    // BMES model groups for Model Info: products are the line shifts without their last "_" part
    var bmesGroups = new List<ModelGroupRecord>
    {
        new() { Name = "SPK", ProductGroup = "SPK", SortOrder = 2, MidGroups = [new() { Material = "SPK-1", SubGroups = [new() { Name = "", LineShifts = ["SPK-1_A"] }] }] },
        new()
        {
            Name = "UNIT", ProductGroup = "UNIT", SortOrder = 1,
            MidGroups =
            [
                new()
                {
                    Material = "TIU-C11-18",
                    SubGroups =
                    [
                        new()
                        {
                            Name = "TIU-C11-18", LineShifts = ["TIU-C11-18-L-ZZ_E21A", "TIU-C11-18-L-ZZ_E22B"],
                            SubGroups = [new() { Name = "TIU-C11-18_L", LineShifts = ["TIU-C11-18-R-ZZ_E21A", "PLAIN", "TIU-C11-18-L-ZZ_E21A"] }],
                        },
                    ],
                },
            ],
        },
    };
    var bmesMaterials = PortalGNLAB.Web.Services.WorkingBmesService.BuildMaterials(bmesGroups);
    var tiu = bmesMaterials.FirstOrDefault(m => m.Material == "TIU-C11-18");
    Check(bmesMaterials.Select(m => m.Group).SequenceEqual(["UNIT", "SPK"]) && tiu is not null
          && tiu.Products.SequenceEqual(["TIU-C11-18-L-ZZ", "TIU-C11-18-R-ZZ", "PLAIN"]),
        "working bmes: products drop the last _ part, keep first-seen order, no duplicates, names without _ stay");
    Check(tiu!.SubGroups.Count == 2 && tiu.SubGroups[0].Path == "TIU-C11-18" && tiu.SubGroups[1].Path == "TIU-C11-18 › TIU-C11-18_L"
          && tiu.SubGroups[1].LineShifts.Count == 3 && bmesMaterials[1].SubGroups[0].Path == "",
        "working bmes: sub-group paths join nested names with ›");

    Check(PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.pdf") == "application/pdf"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.png") == "image/png"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.jpg") == "image/jpeg"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.jpeg") == "image/jpeg"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.gif") == "image/gif"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.tif") == "image/tiff"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.tiff") == "image/tiff",
        "working bmes: drawing Content-Type maps pdf / png / jpg / jpeg / gif / tif / tiff");
    Check(PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("DWG-01.PDF") == "application/pdf"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("x.JpEg") == "image/jpeg"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("noext") == "application/octet-stream"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("a.dwg") == "application/octet-stream"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType("") == "application/octet-stream"
          && PortalGNLAB.Web.Services.WorkingBmesService.DrawingContentType(null) == "application/octet-stream",
        "working bmes: drawing Content-Type ignores case; no or unknown extension is octet-stream");

    // latest successful BOM cache of a product, on a temp DB (never the real cache)
    string bomDb = Path.Combine(Path.GetTempPath(), $"working-bom-{Guid.NewGuid():N}.db");
    try
    {
        var bomCache = new BmesBomCacheService(bomDb);
        BmesBomMaterialCandidate Row(string code, int level) => new() { MaterialCode = code, MaterialName = code + "-N", BomLevel = level, UsageQty = 2.5m, UsageUnit = "EA", HasChildren = level == 0 };
        Check(bomCache.TryLoadLatest("TIU-C11-18-L-ZZ") is null, "working bmes: no cache gives no latest BOM");
        bomCache.Save(new BmesBomCacheQuery("TIU-C11-18-L-ZZ", "3200", new DateTime(2026, 1, 5), 6, 2000), [Row("OLD", 0)]);
        Thread.Sleep(15);
        bomCache.Save(new BmesBomCacheQuery("TIU-C11-18-L-ZZ", "3200", new DateTime(2026, 3, 9), 4, 2000), [Row("NEW", 0), Row("CHILD", 1)]);
        Thread.Sleep(15);
        bomCache.Save(new BmesBomCacheQuery("TIU-C11-18-L-ZZ", "3200", new DateTime(2026, 5, 1), 3, 2000), []);
        var latest = bomCache.TryLoadLatest("tiu-c11-18-l-zz");
        Check(latest is not null && latest.Rows.Select(r => r.MaterialCode).SequenceEqual(["NEW", "CHILD"]) && latest.Query.WorkDate == new DateTime(2026, 3, 9),
            "working bmes: latest BOM is the newest cache with rows, whatever its date or depth, and an empty cache is skipped");
        var bom = PortalGNLAB.Web.Services.WorkingBmesService.ToBom("TIU-C11-18-L-ZZ", "cache", latest!);
        Check(bom.WorkDate == "2026-03-09" && bom.Rows.Count == 2 && bom.Rows[1].Level == 1 && bom.Rows[0].HasChildren && bom.Rows[0].UsageQty == 2.5m,
            "working bmes: BOM response carries work date and rows");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        foreach (string s in new[] { bomDb, bomDb + "-wal", bomDb + "-shm" })
            if (File.Exists(s)) File.Delete(s);
    }
}
finally
{
    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
    foreach (string f in new[] { workDb, glossDb, workGrantDb })
        foreach (string s in new[] { f, f + "-wal", f + "-shm" })
            if (File.Exists(s)) File.Delete(s);
}

// Working: a model name with '/' in a {name} route (the page sends encodeURIComponent(name), so '/' comes as %2F).
{
    var slashApp = Microsoft.AspNetCore.Builder.WebApplication.Create();
    slashApp.Urls.Add("http://127.0.0.1:0");
    Microsoft.AspNetCore.Builder.EndpointRouteBuilderExtensions.MapGet(slashApp, "/m/{name}",
        (string name) => PortalGNLAB.Web.Endpoints.WorkingEndpoints.ModelName(name));
    slashApp.StartAsync().GetAwaiter().GetResult();
    try
    {
        string root = slashApp.Urls.First();
        using var http = new HttpClient();
        string Echo(string name) => http.GetStringAsync($"{root}/m/{Uri.EscapeDataString(name)}").GetAwaiter().GetResult();
        Check(Echo("MSM-X526/X626B") == "MSM-X526/X626B"
              && http.GetStringAsync($"{root}/m/MSM-X526%2fX626B").GetAwaiter().GetResult() == "MSM-X526/X626B"
              && Echo("TIU-C11 18%") == "TIU-C11 18%" && Echo("A%20B") == "A%20B",
            "working route gives a model name with '/' back as it was sent (other escapes decoded once)");
    }
    finally { slashApp.StopAsync().GetAwaiter().GetResult(); }
}

if (failures.Count > 0)
    throw new InvalidOperationException($"Host-cutover checks failed: {string.Join(", ", failures)}");

Console.WriteLine($"PASS all host-cutover checks ({checkCount})");

file sealed class MutableOptionsMonitor<T>(T value) : IOptionsMonitor<T> where T : class
{
    public T CurrentValue { get; } = value;

    public T Get(string? name) => CurrentValue;

    public IDisposable? OnChange(Action<T, string?> listener) => null;
}

file sealed class TestWebHostEnvironment(string contentRoot, string webRoot) : IWebHostEnvironment
{
    public string ApplicationName { get; set; } = "PortalGNLAB.Web.HostCutoverChecks";
    public IFileProvider WebRootFileProvider { get; set; } = new NullFileProvider();
    public string WebRootPath { get; set; } = webRoot;
    public string EnvironmentName { get; set; } = "Verification";
    public string ContentRootPath { get; set; } = contentRoot;
    public IFileProvider ContentRootFileProvider { get; set; } = new NullFileProvider();
}
