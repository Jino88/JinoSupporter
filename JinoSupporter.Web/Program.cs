using System.Security.Claims;
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

// NG Rate settings: singleton (reads/writes ngrate_settings.db)
builder.Services.AddSingleton<NgRateSettingsService>();
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

app.MapRazorComponents<App>()
   .AddInteractiveServerRenderMode();

app.Run();
