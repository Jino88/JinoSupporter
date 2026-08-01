using System.Security.Claims;
using System.Diagnostics;
using System.Text.Json;
using Syncfusion.Blazor;
using JinoSupporter.Web.Components;
using JinoSupporter.Web.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Components.Server.Circuits;
using Microsoft.AspNetCore.StaticFiles;

var builder = WebApplication.CreateBuilder(args);

// ── App paths (centralized) ───────────────────────────────────────────────────
// Loaded BEFORE service registration so the DB / NgRate paths in
// AppPathsConfig become the authoritative defaults for everything below.
var appPathsService = new AppPathsService();
var appPaths = appPathsService.Current;
builder.Configuration["Database:Path"] = appPaths.MainDbPath;
builder.Services.AddSingleton(appPathsService);

// ── Services ──────────────────────────────────────────────────────────────────

builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents()
    .AddHubOptions(options =>
    {
        options.MaximumReceiveMessageSize = 50 * 1024 * 1024; // 50 MB
    });

// Cookie authentication
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(o =>
    {
        o.LoginPath         = "/login";
        o.AccessDeniedPath  = "/login";
        o.ExpireTimeSpan    = TimeSpan.FromDays(30);
        o.SlidingExpiration = true;
    });
builder.Services.AddAuthorizationCore();
builder.Services.AddCascadingAuthenticationState();
builder.Services.AddHttpContextAccessor();

// Claude HTTP client
builder.Services.AddHttpClient<ClaudeService>(client =>
{
    client.BaseAddress = new Uri("https://api.anthropic.com/v1/");
    client.Timeout     = TimeSpan.FromSeconds(420);
});

// OpenAI / Codex API client
builder.Services.AddHttpClient<CodexApiService>(client =>
{
    client.BaseAddress = new Uri("https://api.openai.com/v1/");
    client.Timeout     = TimeSpan.FromSeconds(180);
});

// Syncfusion
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(""); // Community license — works without a key
builder.Services.AddSyncfusionBlazor();

// Singleton DB repository (SQLite file shared across all requests)
builder.Services.AddSingleton<WebRepository>();
builder.Services.AddSingleton<MenuPermissionService>();
builder.Services.AddSingleton<ClaudeUsageScraper>();
builder.Services.AddSingleton<CodexUsageScraper>();
builder.Services.AddSingleton<AiProviderSettingsService>();
builder.Services.AddSingleton<DailyTestExtractionSettingsService>();
builder.Services.AddHostedService<DailyTestDataCliRecoveryService>();
builder.Services.AddSingleton<CurrentProblemAnalysisService>();
builder.Services.AddSingleton<ProcessMaterialMappingService>();
builder.Services.AddScoped<ProcessMaterialNgService>();

// NG Rate settings: singleton (reads/writes ngrate_settings.db)
builder.Services.AddSingleton<NgRateSettingsService>();
builder.Services.AddSingleton<AppActivityLogger>();
// NG Rate: scoped (per-connection HTTP client, progress state)
// Which shell the app renders in (classic MainLayout vs redesigned InstrumentLayout).
// Scoped, so one user switching never changes what anyone else sees.
builder.Services.AddScoped<UiModeService>();
builder.Services.AddScoped<NgRateService>();
// NG Rate report: scoped (reads DB files per request)
builder.Services.AddScoped<NgRateReportService>();
// Worker Status: scoped (per-connection HTTP client)
builder.Services.AddScoped<WorkerStatusService>();
// BMES material master: scoped (per-connection HTTP client)
builder.Services.AddScoped<BmesMaterialService>();
// BMES routing scrape: scoped (per-connection HTTP client)
builder.Services.AddScoped<BmesRoutingScrapeService>();
builder.Services.AddScoped<BmesPdmScrapeService>();
builder.Services.AddSingleton<BmesBomCacheService>();
builder.Services.AddScoped<BmesLpaScrapeService>();
builder.Services.AddScoped<BmesLpaHtmlExportService>();
builder.Services.AddScoped<BmesLpaImageService>();
builder.Services.AddScoped<BmesSettingsSyncService>();
// BMES F-Cost: scoped fetcher + report reader (re-uses NgRateSettingsService for credentials)
builder.Services.AddScoped<FCostService>();
builder.Services.AddScoped<FCostReportService>();
builder.Services.AddScoped<FCostCorePartsService>();
builder.Services.AddScoped<IpgDefectService>();
// BMES report → self-contained static HTML export (memory-friendly iframe rendering)
builder.Services.AddScoped<BmesReportHtmlExportService>();
builder.Services.AddScoped<BmesFcostActualService>();
builder.Services.AddScoped<QrBakoDataService>();
// DAILY REPORT dashboard: singleton so one cached snapshot + one background
// refresh is shared by every connected user.
builder.Services.AddSingleton<BmesDailyReportService>();

// Out-of-process Excel helper driver (DRM-clean + folder-pick).
builder.Services.AddSingleton<ExcelHelperRunner>();
builder.Services.AddSingleton<InputDataTestBatchExtractor>();
builder.Services.AddSingleton<InputDataComPipelineService>();
builder.Services.AddSingleton<MicroSpeakerInputDataService>();
builder.Services.AddSingleton<MicroSpeakerAskEvidenceService>();
builder.Services.AddSingleton<MicroSpeakerReviewCaseService>();

// Connected-users tracking (singleton service + scoped circuit handler)
builder.Services.AddSingleton<ConnectedUsersService>();
builder.Services.AddScoped<UserCircuitHandler>();
builder.Services.AddScoped<CircuitHandler>(sp => sp.GetRequiredService<UserCircuitHandler>());

// Listen on all interfaces. Default 5050 for published/prod runs; dev overrides via ASPNETCORE_URLS
// (set by launchSettings.json or the VS Code launch config).
if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ASPNETCORE_URLS")))
    builder.WebHost.UseUrls("http://*:5050");

// ── Pipeline ──────────────────────────────────────────────────────────────────

var app = builder.Build();

// Seed default admin user if no users exist
var repo = app.Services.GetRequiredService<WebRepository>();
if (repo.GetAllUsers().Count == 0)
    repo.AddUser("admin", AuthService.HashPassword("admin123"), AppRoles.Admin);

// One-shot: import legacy ModelBmes/*.json files into ModelGroups DB.
repo.ImportModelBmesJsonIfNeeded(appPaths.ModelBmesJsonFolder);

var activity = app.Services.GetRequiredService<AppActivityLogger>();
activity.Log(
    "Host",
    $"Process boot pid={Environment.ProcessId}, env={app.Environment.EnvironmentName}, contentRoot={app.Environment.ContentRootPath}, dataRoot={AppStoragePaths.RootDirectory}");

AppDomain.CurrentDomain.UnhandledException += (_, e) =>
{
    activity.Log("Host.Fatal", "UnhandledException terminating=" + e.IsTerminating + " " + e.ExceptionObject);
};

TaskScheduler.UnobservedTaskException += (_, e) =>
{
    activity.Log("Host.Fatal", "UnobservedTaskException " + e.Exception);
};

app.Lifetime.ApplicationStarted.Register(() =>
    activity.Log("Host", "ApplicationStarted urls=" + string.Join(", ", app.Urls)));
app.Lifetime.ApplicationStopping.Register(() =>
    activity.Log("Host", "ApplicationStopping"));
app.Lifetime.ApplicationStopped.Register(() =>
    activity.Log("Host", "ApplicationStopped"));

if (!app.Environment.IsDevelopment())
    app.UseExceptionHandler("/Error", createScopeForErrors: true);

app.UseStaticFiles();
app.UseAuthentication();
app.UseAuthorization();
app.Use(async (ctx, next) =>
{
    string path = ctx.Request.Path.Value ?? "";
    bool skipLog =
        path.StartsWith("/_framework", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/_blazor", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/css", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/js", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/lib", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".ico", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".css", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".js", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".map", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
        path.EndsWith(".svg", StringComparison.OrdinalIgnoreCase);

    var sw = Stopwatch.StartNew();
    try
    {
        await next();
    }
    catch (Exception ex)
    {
        sw.Stop();
        string query = ctx.Request.QueryString.HasValue ? ctx.Request.QueryString.Value ?? "" : "";
        ctx.RequestServices.GetRequiredService<AppActivityLogger>()
            .Log("Web.Error", $"{ctx.Request.Method} {path}{query} failed ({sw.ElapsedMilliseconds:N0} ms) {ex}");
        throw;
    }
    finally
    {
        if (sw.IsRunning)
            sw.Stop();

        if (!skipLog)
        {
            string query = ctx.Request.QueryString.HasValue ? ctx.Request.QueryString.Value ?? "" : "";
            ctx.RequestServices.GetRequiredService<AppActivityLogger>()
                .Log("Web", $"{ctx.Request.Method} {path}{query} -> {ctx.Response.StatusCode} ({sw.ElapsedMilliseconds:N0} ms)");
        }
    }
});
app.UseAntiforgery();

// ── Auth endpoints ────────────────────────────────────────────────────────────

app.MapPost("/auth/login", async (HttpContext ctx) =>
{
    var form     = await ctx.Request.ReadFormAsync();
    string user  = form["username"].ToString().Trim();
    string pass  = form["password"].ToString();
    string ret   = form["returnUrl"].ToString();
    if (string.IsNullOrWhiteSpace(ret) || !ret.StartsWith('/')) ret = "/";

    var record = repo.GetUser(user);
    if (record is null || !AuthService.VerifyPassword(pass, record.PasswordHash))
    {
        ctx.Response.Redirect($"/login?error=1&returnUrl={Uri.EscapeDataString(ret)}");
        return;
    }

    var claims    = new[] { new Claim(ClaimTypes.Name, record.Username), new Claim(ClaimTypes.Role, record.Role) };
    var identity  = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
    var principal = new ClaimsPrincipal(identity);
    await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal);
    ctx.Response.Redirect(ret);
}).DisableAntiforgery();

app.MapPost("/auth/logout", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    ctx.Response.Redirect("/login");
}).DisableAntiforgery();

// ── Raw file download (DataInference attachments) ─────────────────────────────

app.MapGet("/data-inference/file/{id:long}", (long id, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    var file = repo.GetRawReportFile(id);
    if (file is null) return Results.NotFound();
    string safeName = string.IsNullOrEmpty(file.Value.FileName) ? $"file-{id}" : file.Value.FileName;
    return Results.File(file.Value.Data, file.Value.MediaType, safeName);
});

// Export full dataset bundle (images + measurements + summary + issues) as a ZIP,
// intended for feeding back to Claude / external review.
app.MapGet("/data-inference/export/{datasetName}", (string datasetName, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    byte[] zip = DatasetExportBuilder.BuildZip(repo, datasetName);
    string safeName = System.Text.RegularExpressions.Regex.Replace(datasetName, @"[^\w\-.]+", "_");
    return Results.File(zip, "application/zip", $"{safeName}.zip");
});

// Export all datasets that have validation issues as one ZIP (each nested as /<name>/...).
app.MapGet("/data-inference/export-all", (HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    byte[] zip = DatasetExportBuilder.BuildAllFlaggedZip(repo);
    string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    return Results.File(zip, "application/zip", $"flagged_datasets_{ts}.zip");
});

// Serve the generated BMES report as a single self-contained static HTML file (all
// tabs + menu inside it), shown in an iframe on /report/bmes.
app.MapGet("/report/bmes/view/{token}", (string token, BmesReportHtmlExportService svc, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    string? path = svc.ResolveReportFile(token);
    if (path is null) return Results.NotFound();
    return Results.File(path, "text/html; charset=utf-8");
});

// Serve the generated LPA result as a single self-contained static HTML file (list +
// NG pivot + every row's checklist inside it), shown in an iframe on /bmes/lpa.
app.MapGet("/bmes/lpa/view/{token}", (string token, BmesLpaHtmlExportService svc, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    string? path = svc.ResolveReportFile(token);
    if (path is null) return Results.NotFound();
    return Results.File(path, "text/html; charset=utf-8");
});

// One LPA result photo, downscaled and cached in the DB, fetched on demand by the viewer
// (size=thumb|view) instead of being embedded — so Search no longer blocks on downloading
// every photo, and each one leaves BMES exactly once.
app.MapGet("/bmes/lpa/img", async (string? path, string? size, BmesLpaImageService svc, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();
    if (string.IsNullOrWhiteSpace(path)) return Results.BadRequest();
    byte[]? bytes = await svc.GetAsync(path, size == "view");
    if (bytes is null) return Results.NotFound();
    // Path → bytes is stable (a photo is one upload slot of one closed audit), so let the
    // browser keep it and not re-ask while scrolling the same table.
    ctx.Response.Headers.CacheControl = "private, max-age=31536000, immutable";
    return Results.File(bytes, "image/jpeg");
});

app.MapGet("/microspeaker/dashboard/{kind}", (string kind, MicroSpeakerInputDataService microSpeaker, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();

    MicroSpeakerPaths paths = microSpeaker.ResolvePaths();
    string dashboardPath = kind.Trim().ToLowerInvariant() switch
    {
        "input" or "cli" => paths.DashboardPath,
        "pair" or "compare" => Path.Combine(paths.ProjectRoot, "db", "report_compare_dashboard.html"),
        "pair-review" => Path.Combine(paths.ProjectRoot, "db", "p0_rule_gap_pair_review.html"),
        _ => ""
    };

    if (string.IsNullOrWhiteSpace(dashboardPath))
        return Results.BadRequest(new { message = "Unknown MicroSpeaker dashboard." });

    dashboardPath = Path.GetFullPath(dashboardPath);
    if (!File.Exists(dashboardPath))
        return Results.NotFound(new { message = "MicroSpeaker dashboard was not found.", path = dashboardPath });

    return Results.File(dashboardPath, "text/html; charset=utf-8");
});

app.MapGet("/microspeaker/source-file/{fileId:long}", (long fileId, MicroSpeakerInputDataService microSpeaker, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();

    MicroSpeakerSourceFile? source = microSpeaker.FindSourceFile(fileId);
    if (source is null)
        return Results.NotFound(new { message = "MicroSpeaker source file was not found.", fileId });

    if (!File.Exists(source.FullPath))
        return Results.NotFound(new { message = "MicroSpeaker source file path does not exist.", fileId, path = source.FullPath });

    var provider = new FileExtensionContentTypeProvider();
    if (!provider.TryGetContentType(source.FullPath, out string? contentType))
        contentType = "application/octet-stream";

    string downloadName = string.IsNullOrWhiteSpace(source.FileName)
        ? Path.GetFileName(source.FullPath)
        : source.FileName;
    return Results.File(source.FullPath, contentType, downloadName, enableRangeProcessing: true);
});

app.MapGet("/microspeaker/review-cases/sample.json", (int? limit, string? q, MicroSpeakerReviewCaseService reviewCases, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();

    MicroSpeakerReviewCaseSample sample = reviewCases.BuildSample(limit ?? 80, q);
    return Results.Json(sample, new JsonSerializerOptions
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    });
});

app.MapGet("/microspeaker/review-cases/ai-packet/{fileId:long}.json", (long fileId, int? rowLimit, int? candidateLimit, MicroSpeakerReviewCaseService reviewCases, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();

    MicroSpeakerReviewCaseAiPacket packet = reviewCases.BuildAiPacket(fileId, rowLimit ?? 1200, candidateLimit ?? 300);
    if (packet.DatabaseExists && !packet.FileFound)
        return Results.NotFound(packet);

    return Results.Json(packet, new JsonSerializerOptions
    {
        WriteIndented = true,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
    });
});

// ── Blazor ────────────────────────────────────────────────────────────────────

// Standalone app update feed. Packages are generated by
// BmesNgRateStandalone/tools/PublishStandaloneUpdate.ps1 into:
// JinoSupporter.Web/standalone-updates/
app.MapGet("/data-inference/ask-history/{id:long}/html", (long id, string? lang, WebRepository repo, HttpContext ctx) =>
{
    if (ctx.User?.Identity?.IsAuthenticated != true) return Results.Unauthorized();

    AskAiHistoryRecord? history = repo.GetAskAiHistoryById(id);
    if (history is null)
        return Results.NotFound(new { message = "Ask AI history was not found.", id });

    string html = AskHistoryOverallHtml(history, lang);
    return Results.Content(PrepareAskHistoryHtml(html), "text/html; charset=utf-8");
});

string standaloneUpdateDir = Path.Combine(app.Environment.ContentRootPath, "standalone-updates");

app.MapGet("/standalone/update.json", () =>
{
    string manifestPath = Path.Combine(standaloneUpdateDir, "update.json");
    return File.Exists(manifestPath)
        ? Results.File(manifestPath, "application/json")
        : Results.NotFound(new { message = "No standalone update manifest has been published." });
});

// Serves both the .zip the in-app updater pulls and the .exe installer the PC Download
// page hands out for a fresh machine. Anonymous on purpose: the updater runs before login.
app.MapGet("/standalone/download/{fileName}", (string fileName) =>
{
    string safeName = Path.GetFileName(fileName);
    bool isZip = safeName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase);
    bool isExe = safeName.EndsWith(".exe", StringComparison.OrdinalIgnoreCase);
    if (!isZip && !isExe)
        return Results.BadRequest("Only zip update packages and exe installers are allowed.");

    string packagePath = Path.Combine(standaloneUpdateDir, safeName);
    return File.Exists(packagePath)
        ? Results.File(packagePath, isZip ? "application/zip" : "application/octet-stream", safeName)
        : Results.NotFound();
});

app.MapPost("/standalone/sync/routing-table", (NgRateSettingsService settings, List<RoutingRow> rows) =>
{
    settings.ReplaceRoutingRows(rows);
    return Results.Ok(new { rows = rows.Count });
});

app.MapGet("/standalone/sync/routing-table", (NgRateSettingsService settings) =>
{
    var rows = settings.GetRoutingRows();
    return Results.Ok(rows);
});

app.MapPost("/standalone/sync/reason-table", (NgRateSettingsService settings, List<ReasonRow> rows) =>
{
    settings.ReplaceReasonRows(rows);
    return Results.Ok(new { rows = rows.Count });
});

app.MapGet("/standalone/sync/reason-table", (NgRateSettingsService settings) =>
{
    var rows = settings.GetReasonRows();
    return Results.Ok(rows);
});

app.MapPost("/standalone/sync/model-groups", (WebRepository repo, List<ModelGroupRecord> groups) =>
{
    repo.SaveModelGroups(groups);
    return Results.Ok(new { rows = groups.Count });
});

app.MapGet("/standalone/sync/model-groups", (WebRepository repo) =>
{
    var groups = repo.GetModelGroups();
    return Results.Ok(groups);
});

app.MapPost("/standalone/sync/bmes-materials", (WebRepository repo, List<BmesMaterial> rows) =>
{
    int saved = repo.UpsertBmesMaterials(rows);
    return Results.Ok(new { rows = saved });
});

app.MapGet("/standalone/sync/bmes-materials", (WebRepository repo) =>
{
    var rows = repo.GetBmesMaterials();
    return Results.Ok(rows);
});

static string AskHistoryOverallHtml(AskAiHistoryRecord history, string? lang)
{
    string code = (lang ?? "").Trim().ToLowerInvariant();
    if (code is "ko" or "en" or "vi" && !string.IsNullOrWhiteSpace(history.TranslationsJson))
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(history.TranslationsJson);
            if (doc.RootElement.TryGetProperty(code, out JsonElement translated)
                && translated.TryGetProperty("overall", out JsonElement overall)
                && overall.ValueKind == JsonValueKind.String)
            {
                string? value = overall.GetString();
                if (!string.IsNullOrWhiteSpace(value)) return value;
            }
        }
        catch (JsonException)
        {
            // Fall back to the stored main HTML.
        }
    }

    return history.Overall;
}

static string PrepareAskHistoryHtml(string html)
{
    string text = html ?? "";
    if (text.Contains("jino-ask-ai-report-css", StringComparison.OrdinalIgnoreCase))
        return text;

    string css = string.Join(Environment.NewLine, new[]
    {
        "<style id=\"jino-ask-ai-report-css\">",
        "html, body { margin: 0; padding: 10px; color: #0f172a; background: #f8fafc; font-family: \"Malgun Gothic\", \"Segoe UI\", Arial, sans-serif; font-size: 12px; line-height: 1.45; }",
        "* { box-sizing: border-box; }",
        "h1 { margin: 0 0 10px; font-size: 18px; line-height: 1.25; }",
        "h2 { margin: 18px 0 8px; font-size: 15px; line-height: 1.3; }",
        "h3 { margin: 14px 0 6px; font-size: 13px; line-height: 1.3; }",
        "p { margin: 4px 0 8px; }",
        "table { width: 100%; min-width: 1120px; border-collapse: collapse; table-layout: auto; background: #fff; font-size: 11px; }",
        "th, td { border: 1px solid #cbd5e1; padding: 6px 7px; text-align: left !important; vertical-align: top !important; white-space: normal !important; word-break: keep-all; overflow-wrap: break-word; }",
        "th { background: #e2e8f0; color: #1e293b; font-weight: 700; }",
        "td { color: #0f172a; }",
        "a[href^=\"/microspeaker/source-file/\"] { display: inline-block; white-space: nowrap; font-weight: 700; color: #1d4ed8; text-decoration: none; }",
        "a[href^=\"/microspeaker/source-file/\"]:hover { text-decoration: underline; }",
        ".num, .rate, .pct, td[data-type=\"number\"] { text-align: right !important; font-variant-numeric: tabular-nums; white-space: nowrap !important; }",
        ".limit, .judgement { max-width: 180px; }",
        "svg { max-width: 100%; }",
        "</style>",
    });

    int headClose = text.IndexOf("</head>", StringComparison.OrdinalIgnoreCase);
    if (headClose >= 0) return text.Insert(headClose, css);

    int bodyOpen = text.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
    if (bodyOpen >= 0)
    {
        int bodyEnd = text.IndexOf('>', bodyOpen);
        if (bodyEnd >= 0) return text.Insert(bodyEnd + 1, css);
    }

    return css + text;
}

app.MapRazorComponents<App>()
   .AddInteractiveServerRenderMode();

try
{
    app.Run();
}
catch (Exception ex)
{
    activity.Log("Host.Fatal", "Run failed " + ex);
    throw;
}
