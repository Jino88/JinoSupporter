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

// Production Input: models, days, plan, cascade delete, Excel export and the menu entry.
string prodDb = Path.Combine(Path.GetTempPath(), $"prod-input-{Guid.NewGuid():N}.db");
try
{
    var prod = new ProductionInputService(prodDb);
    JsonElement J(string json) => JsonDocument.Parse(json).RootElement.Clone();
    ProductionModelSave ModelBody(string name, string build, string setup, int rev = 0) => new(name, build, J(setup), rev);
    bool Conflicts(Action act)
    {
        try { act(); return false; }
        catch (ProductionConflictException) { return true; }
    }
    bool BadRequest(Action act)
    {
        try { act(); return false; }
        catch (ArgumentException) { return true; }
    }

    ProductionModel x316 = prod.CreateModel(ModelBody("  X316B ", "EVT", """{"processes":[{"n":"A"},{"n":"B"}],"materials":[{"n":"m"}],"extra":{"keep":1}}"""), "tester");
    Check(x316.Id.Length == 32 && x316.Name == "X316B" && x316.Revision == 1 && x316.UpdatedBy == "tester"
          && x316.Setup.GetProperty("extra").GetProperty("keep").GetInt32() == 1,
        "production model create trims the name, starts at revision 1 and keeps setup verbatim");
    Check(Conflicts(() => prod.CreateModel(ModelBody("x316b", "evt", "{}"), "tester")),
        "production model with the same name and build is a conflict (case-insensitive)");
    Check(BadRequest(() => prod.CreateModel(ModelBody("   ", "", "{}"), "t"))
          && BadRequest(() => prod.CreateModel(ModelBody(new string('n', 81), "", "{}"), "t"))
          && BadRequest(() => prod.CreateModel(ModelBody("n", new string('b', 41), "{}"), "t")),
        "production model name is required and length limited");
    ProductionModel a100 = prod.CreateModel(ModelBody("A100", "", "{}"), "tester");
    List<ProductionModelSummary> listed = prod.ListModels();
    Check(listed.Select(m => m.Name).SequenceEqual(["A100", "X316B"])
          && listed[1].ProcessCount == 2 && listed[1].MaterialCount == 1 && listed[0].ProcessCount == 0,
        "production model list is ordered by name and counts setup processes and materials");

    ProductionModel? updated = prod.UpdateModel(x316.Id, ModelBody("X316B", "EVT", """{"processes":[]}""", 1), "editor");
    Check(updated is { Revision: 2, UpdatedBy: "editor" }, "production model update bumps the revision");
    Check(Conflicts(() => prod.UpdateModel(x316.Id, ModelBody("X316B", "EVT", "{}", 1), "editor")),
        "production model update with a stale revision is a conflict");
    Check(Conflicts(() => prod.UpdateModel(a100.Id, ModelBody("X316B", "EVT", "{}", 1), "editor")),
        "production model rename onto an existing name and build is a conflict");
    Check(prod.UpdateModel("missing", ModelBody("Z", "", "{}", 1), "e") is null && prod.GetModel("missing") is null,
        "production model lookups of an unknown id return nothing");

    ProductionDay? d1 = prod.SaveDay(x316.Id, "2026-10-01", new("Draft", J("""{"a":1}"""), J("""{"s":2}"""), 0), "alice");
    ProductionDay? d3 = prod.SaveDay(x316.Id, "2026-10-03", new("Submitted", J("[1,2]"), J("{}"), 0), "bob");
    Check(d1 is { Revision: 1, EnteredBy: "alice", Status: "Draft" } && d1.Data.GetProperty("a").GetInt32() == 1 && d3 is { Revision: 1 },
        "production day create stores data verbatim and the signed-in user");
    Check(Conflicts(() => prod.SaveDay(x316.Id, "2026-10-01", new("Draft", J("{}"), J("{}"), 0), "x")),
        "production day create over an existing day is a conflict");
    Check(Conflicts(() => prod.SaveDay(x316.Id, "2026-10-01", new("Draft", J("{}"), J("{}"), 5), "x")),
        "production day update with a stale revision is a conflict");
    ProductionDay? d1b = prod.SaveDay(x316.Id, "2026-10-01", new("Submitted", J("""{"a":9}"""), J("{}"), 1), "carol");
    Check(d1b is { Revision: 2, Status: "Submitted", EnteredBy: "carol" }, "production day update bumps the revision");
    Check(BadRequest(() => prod.SaveDay(x316.Id, "2026-10-01", new("Done", J("{}"), J("{}"), 2), "x"))
          && BadRequest(() => prod.SaveDay(x316.Id, "10/01/2026", new("Draft", J("{}"), J("{}"), 0), "x"))
          && prod.SaveDay("missing", "2026-10-01", new("Draft", J("{}"), J("{}"), 0), "x") is null,
        "production day rejects a bad status or date and an unknown model");
    ProductionDayPair pair = prod.GetDay(x316.Id, "2026-10-02");
    ProductionDayPair pairOnDay = prod.GetDay(x316.Id, "2026-10-03");
    Check(pair.Day is null && pair.Previous?.Date == "2026-10-01"
          && pairOnDay.Day?.Date == "2026-10-03" && pairOnDay.Previous?.Date == "2026-10-01"
          && prod.GetDay(x316.Id, "2026-09-30").Previous is null,
        "production previous day is the latest stored day before the date");
    Check(BadRequest(() => prod.GetDay(x316.Id, "2026-13-01")), "production day lookup rejects a bad date");
    Check(prod.ListDays(x316.Id, "2026-10-01", "2026-10-03").Select(d => d.Date).SequenceEqual(["2026-10-01", "2026-10-03"])
          && prod.ListDays(x316.Id, "2026-10-02", "2026-10-03").Count == 1
          && prod.ListDays(x316.Id, "2026-10-03", "2026-10-03").Count == 1,
        "production day range is ascending and inclusive");
    Check(BadRequest(() => prod.ListDays(x316.Id, "2026-10-01", null))
          && BadRequest(() => prod.ListDays(x316.Id, "2025-01-01", "2026-10-03"))
          && prod.ListDays(x316.Id, "2025-08-30", "2026-10-03").Count == 2,
        "production day range needs both ends and spans at most 400 days");

    Check(prod.GetPlan(x316.Id) is { Data: null, Revision: 0 }, "production plan is empty with revision 0 before the first save");
    ProductionPlan? plan1 = prod.SavePlan(x316.Id, new(J("""{"w1":100}"""), 0), "pm");
    Check(plan1 is { Revision: 1, UpdatedBy: "pm" } && plan1.Data?.GetProperty("w1").GetInt32() == 100, "production plan create stores data verbatim");
    Check(Conflicts(() => prod.SavePlan(x316.Id, new(J("{}"), 0), "pm")) && Conflicts(() => prod.SavePlan(x316.Id, new(J("{}"), 7), "pm")),
        "production plan save with a wrong revision is a conflict");
    Check(prod.SavePlan(x316.Id, new(J("""{"w1":200}"""), 1), "pm") is { Revision: 2 }, "production plan update bumps the revision");

    prod.DeleteDay(x316.Id, "2026-10-03");
    Check(prod.GetDay(x316.Id, "2026-10-03").Day is null, "production day delete removes the entry");
    Check(prod.DeleteModel(x316.Id) && !prod.DeleteModel(x316.Id)
          && prod.ListDays(x316.Id, "2026-10-01", "2026-10-31").Count == 0 && prod.GetPlan(x316.Id).Revision == 0
          && prod.GetModel(a100.Id) is not null,
        "production model delete also removes its days and plan");

    byte[] xlsx = ProductionInputService.ExportXlsx(new ExportRequest("report", [
        new ExportSheet("Sheet: A/B", [
            [J("\"Head\""), J("\"\""), J("\"Total\"")],
            [J("\"x\""), J("12.5"), J("null")],
        ], [new ExportMerge(0, 0, 1, 2)], 1, [20, 10]),
        new ExportSheet("Sheet: A/B", [[J("1")]], null, 0, null),
    ]));
    using (var book = new ClosedXML.Excel.XLWorkbook(new MemoryStream(xlsx)))
    {
        var ws = book.Worksheet(1);
        Check(book.Worksheets.Count == 2 && ws.Name == "Sheet AB" && book.Worksheet(2).Name == "Sheet AB (2)"
              && ws.MergedRanges.Single().RangeAddress.ToStringRelative() == "A1:B1"
              && ws.Cell(2, 2).DataType == ClosedXML.Excel.XLDataType.Number && ws.Cell(2, 2).GetDouble() == 12.5
              && ws.Cell(1, 1).Style.Font.Bold && ws.Cell(2, 1).Style.Font.Bold == false
              && ws.SheetView.SplitRow == 1 && ws.Column(1).Width == 20
              && ws.Cell(2, 1).Style.Border.TopBorder == ClosedXML.Excel.XLBorderStyleValues.Thin,
            "production export keeps merges, numeric cells, header style, freeze, widths and unique sheet names");
    }
    Check(ProductionInputService.SafeFileName("..\\a:b/ report?.XLSX") == "report_.xlsx" && ProductionInputService.SafeFileName("  ") == "production-input.xlsx",
        "production export file name is sanitized and ends with .xlsx");
    Check(BadRequest(() => ProductionInputService.ExportXlsx(new ExportRequest("x", Enumerable.Range(0, 21).Select(i => new ExportSheet("s" + i, [[J("1")]], null, 0, null)).ToList())))
          && BadRequest(() => ProductionInputService.ExportXlsx(new ExportRequest("x", [new ExportSheet("big", Enumerable.Range(0, 201).Select(_ => Enumerable.Range(0, 1000).Select(_ => J("1")).ToList()).ToList(), null, 0, null)]))),
        "production export rejects more than 20 sheets or 200k cells");

    var managerMenu = PortalGNLAB.Web.Endpoints.PortalShellEndpoints.MenuFor(AppRoles.Manager, (role, id) => AppMenus.DefaultsByRole[role].Contains(id));
    var viewerMenu = PortalGNLAB.Web.Endpoints.PortalShellEndpoints.MenuFor(AppRoles.Viewer, (role, id) => AppMenus.DefaultsByRole[role].Contains(id));
    Check(managerMenu.Any(m => m.Id == AppMenus.ProductionInput && m.React && m.Href == "/app/production/input" && m.Group == "Production" && m.Label == "Production Input")
          && !viewerMenu.Any(m => m.Id == AppMenus.ProductionInput),
        "production-input menu is offered to Manager with the React href and not to Viewer");

    // Existing installs: roles already seeded get the menu once, and an admin removal sticks afterwards.
    string grantDb = Path.Combine(Path.GetTempPath(), $"prod-grant-{Guid.NewGuid():N}.db");
    try
    {
        var grantRepo = new WebRepository(Microsoft.Extensions.Configuration.MemoryConfigurationBuilderExtensions.AddInMemoryCollection(
            new Microsoft.Extensions.Configuration.ConfigurationBuilder(),
            new Dictionary<string, string?> { ["Database:Path"] = grantDb }).Build());
        grantRepo.SetMenuPermission(AppRoles.Leader, AppMenus.NgRate, true);
        var perms = new MenuPermissionService(grantRepo);
        bool granted = perms.IsAllowed(AppRoles.Leader, AppMenus.ProductionInput);
        grantRepo.SetMenuPermission(AppRoles.Leader, AppMenus.ProductionInput, false);
        _ = new MenuPermissionService(grantRepo);
        Check(granted && !perms.IsAllowed(AppRoles.Leader, AppMenus.ProductionInput)
              && grantRepo.GetMenuPermissionsForRole(AppRoles.Viewer).Contains(AppMenus.ProductionInput) == false,
            "production-input is granted once to roles seeded before it existed");
    }
    finally
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        foreach (string f in new[] { grantDb, grantDb + "-wal", grantDb + "-shm" })
            if (File.Exists(f)) File.Delete(f);
    }
}
finally
{
    Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
    foreach (string f in new[] { prodDb, prodDb + "-wal", prodDb + "-shm" })
        if (File.Exists(f)) File.Delete(f);
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
