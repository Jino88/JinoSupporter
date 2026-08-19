using System.Text.Json;
using JinoSupporter.Web.Services;
using JinoSupporter.Web.Services.BmesReports;
using JinoSupporter.Web.Services.BmesReports.Contracts;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;

if (args.Length != 1)
    throw new ArgumentException("Pass the absolute worktree root as the only argument.");

string workspace = Path.GetFullPath(args[0]);
string webProject = Path.Combine(workspace, "JinoSupporter.Web");
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
    public string ApplicationName { get; set; } = "JinoSupporter.Web.HostCutoverChecks";
    public IFileProvider WebRootFileProvider { get; set; } = new NullFileProvider();
    public string WebRootPath { get; set; } = webRoot;
    public string EnvironmentName { get; set; } = "Verification";
    public string ContentRootPath { get; set; } = contentRoot;
    public IFileProvider ContentRootFileProvider { get; set; } = new NullFileProvider();
}
