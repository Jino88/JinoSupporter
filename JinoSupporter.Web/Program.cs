using System.Security.Claims;
using System.Diagnostics;
using Syncfusion.Blazor;
using JinoSupporter.Web.Components;
using JinoSupporter.Web.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Components.Server.Circuits;

var builder = WebApplication.CreateBuilder(args);

// ── App paths (centralized) ───────────────────────────────────────────────────
// Loaded BEFORE service registration so the DB / NgRate paths in
// AppPathsConfig become the authoritative defaults for everything below.
var appPathsService = new AppPathsService();
var appPaths = appPathsService.Current;
builder.Configuration["Database:Path"] = appPaths.MainDbPath;
builder.Configuration["Schedule:Path"] = appPaths.ScheduleDbPath;
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

// Syncfusion
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(""); // Community license — works without a key
builder.Services.AddSyncfusionBlazor();

// Singleton DB repository (SQLite file shared across all requests)
builder.Services.AddSingleton<WebRepository>();
builder.Services.AddSingleton<MenuPermissionService>();
builder.Services.AddSingleton<ClaudeUsageScraper>();
builder.Services.AddSingleton<AiProviderSettingsService>();

// NG Rate settings: singleton (reads/writes ngrate_settings.db)
builder.Services.AddSingleton<NgRateSettingsService>();
builder.Services.AddSingleton<AppActivityLogger>();
// NG Rate: scoped (per-connection HTTP client, progress state)
builder.Services.AddScoped<NgRateService>();
// NG Rate report: scoped (reads DB files per request)
builder.Services.AddScoped<NgRateReportService>();
// Worker Status: scoped (per-connection HTTP client)
builder.Services.AddScoped<WorkerStatusService>();
// BMES material master: scoped (per-connection HTTP client)
builder.Services.AddScoped<BmesMaterialService>();
// BMES routing scrape: scoped (per-connection HTTP client)
builder.Services.AddScoped<BmesRoutingScrapeService>();
builder.Services.AddScoped<BmesSettingsSyncService>();
// BMES F-Cost: scoped fetcher + report reader (re-uses NgRateSettingsService for credentials)
builder.Services.AddScoped<FCostService>();
builder.Services.AddScoped<FCostReportService>();

// Out-of-process Excel helper driver (DRM-clean + folder-pick).
builder.Services.AddSingleton<ExcelHelperRunner>();

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
    await next();
    sw.Stop();

    if (!skipLog)
    {
        string query = ctx.Request.QueryString.HasValue ? ctx.Request.QueryString.Value ?? "" : "";
        ctx.RequestServices.GetRequiredService<AppActivityLogger>()
            .Log("Web", $"{ctx.Request.Method} {path}{query} -> {ctx.Response.StatusCode} ({sw.ElapsedMilliseconds:N0} ms)");
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

// ── Blazor ────────────────────────────────────────────────────────────────────

// Standalone app update feed. Packages are generated by
// BmesNgRateStandalone/tools/PublishStandaloneUpdate.ps1 into:
// JinoSupporter.Web/standalone-updates/
string standaloneUpdateDir = Path.Combine(app.Environment.ContentRootPath, "standalone-updates");

app.MapGet("/standalone/update.json", () =>
{
    string manifestPath = Path.Combine(standaloneUpdateDir, "update.json");
    return File.Exists(manifestPath)
        ? Results.File(manifestPath, "application/json")
        : Results.NotFound(new { message = "No standalone update manifest has been published." });
});

app.MapGet("/standalone/download/{fileName}", (string fileName) =>
{
    string safeName = Path.GetFileName(fileName);
    if (!safeName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        return Results.BadRequest("Only zip update packages are allowed.");

    string packagePath = Path.Combine(standaloneUpdateDir, safeName);
    return File.Exists(packagePath)
        ? Results.File(packagePath, "application/zip", safeName)
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

app.MapRazorComponents<App>()
   .AddInteractiveServerRenderMode();

app.Run();
