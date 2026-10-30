using System.Security.Claims;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// JSON API and hosting for the React app (ClientApp/Portal, built into wwwroot/app).
/// The React app lives under /app while the Blazor pages are moved over one by one: the
/// menu lists every page the user may open and points either at the old Blazor route or,
/// once a page is ported, at its /app route (<see cref="ReactRoutes"/>).
/// </summary>
public static class PortalShellEndpoints
{
    /// <summary>Blazor route of every menu (same targets as Components/Layout/NavMenu.razor).</summary>
    public static readonly IReadOnlyDictionary<string, string> BlazorRoutes = new Dictionary<string, string>
    {
        [AppMenus.NgRate]             = "/ng-rate",
        [AppMenus.BmesFCost]          = "/bmes/f-cost",
        [AppMenus.BmesWorkerStatus]   = "/bmes/worker-status",
        [AppMenus.BmesMakeModelGroup] = "/bmes/make-model-group",
        [AppMenus.BmesRoutingTable]   = "/bmes/routing-table",
        [AppMenus.BmesReasonTable]    = "/bmes/reason-table",
        [AppMenus.BmesTest5]          = "/bmes/test5",
        [AppMenus.BmesDailyReport]    = "/bmes/daily-report",
        [AppMenus.QrSummary]          = "/qr/summary",
        [AppMenus.QrBakoData]         = "/qr/bako-data",
        [AppMenus.QrAirleakData]      = "/qr/airleak-data",
        [AppMenus.QrGaussData]        = "/qr/gauss-data",
        [AppMenus.Report]             = "/report/bmes",
        [AppMenus.Translate]          = "/translate",
        [AppMenus.ReviewReport]       = "/tools/review-report",
        [AppMenus.ProductionInput]    = "/app/production/input",
        [AppMenus.BmesSetting]        = "/setting",
        [AppMenus.AdminUsers]         = "/admin/users",
        [AppMenus.AdminAiUsages]      = "/admin/ai-usages",
        [AppMenus.TestExcelConverter] = "/admin/test-excel-converter",
        [AppMenus.AdminDbQuery]       = "/admin/db-query",
    };

    /// <summary>Menus already ported to React, with their route inside the React app.</summary>
    public static readonly IReadOnlyDictionary<string, string> ReactRoutes = new Dictionary<string, string>
    {
        [AppMenus.AdminUsers] = "/app/admin/users",
        [AppMenus.AdminAiUsages] = "/app/admin/ai-usages",
        [AppMenus.AdminDbQuery] = "/app/admin/db-query",
        [AppMenus.TestExcelConverter] = "/app/admin/test-excel-converter",
        [AppMenus.Translate] = "/app/tools/translate",
        [AppMenus.ReviewReport] = "/app/tools/review-report",
        [AppMenus.ProductionInput] = "/app/production/input",
        [AppMenus.BmesRoutingTable] = "/app/bmes/routing-table",
        [AppMenus.BmesReasonTable] = "/app/bmes/reason-table",
        [AppMenus.BmesWorkerStatus] = "/app/bmes/worker-status",
        [AppMenus.BmesMakeModelGroup] = "/app/bmes/make-model-group",
        [AppMenus.QrSummary] = "/app/qr/summary",
        [AppMenus.QrBakoData] = "/app/qr/bako-data",
        [AppMenus.QrAirleakData] = "/app/qr/airleak-data",
        [AppMenus.QrGaussData] = "/app/qr/gauss-data",
    };

    public sealed record PortalUser(string Name, string Role);

    public sealed record PortalMenuItem(string Id, string Label, string Group, string Href, bool React);

    public sealed record PortalDownloadAsset(string FileName, long SizeBytes, string SizeText, string Url);

    public sealed record PortalDownload(
        string Version, string Notes, DateTimeOffset? PublishedAt, PortalDownloadAsset? Setup, PortalDownloadAsset? Package);

    /// <summary>Menu entries the role may open, in <see cref="AppMenus.All"/> order. Menus with no
    /// page (legacy permission keys) are left out.</summary>
    public static List<PortalMenuItem> MenuFor(string role, Func<string, string, bool> isAllowed) =>
        AppMenus.All
            .Where(m => BlazorRoutes.ContainsKey(m.Id) && isAllowed(role, m.Id))
            .Select(m => ReactRoutes.TryGetValue(m.Id, out string? react)
                ? new PortalMenuItem(m.Id, m.Label, m.Group, react, true)
                : new PortalMenuItem(m.Id, m.Label, m.Group, BlazorRoutes[m.Id], false))
            .ToList();

    public static PortalDownload? ToPortalDownload(StandaloneRelease? release)
    {
        if (release is null) return null;
        static PortalDownloadAsset? Asset(StandaloneAsset? a) =>
            a is null ? null : new PortalDownloadAsset(a.FileName, a.SizeBytes, a.SizeText, a.Url);
        return new PortalDownload(release.Version, release.Notes, release.PublishedAt, Asset(release.Setup), Asset(release.Package));
    }

    /// <summary>
    /// index.html names the hashed bundle of the current build, so the browser must ask again on
    /// every load; without this it may keep an old index.html and show the screen from before a restart.
    /// </summary>
    public static readonly StaticFileOptions ShellFileOptions = new()
    {
        OnPrepareResponse = file => file.Context.Response.Headers.CacheControl = "no-cache, no-store, must-revalidate",
    };

    public static IEndpointRouteBuilder MapPortalShell(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal").RequireAuthorization();

        api.MapGet("/me", (ClaimsPrincipal user) =>
            new PortalUser(user.Identity?.Name ?? "", user.FindFirst(ClaimTypes.Role)?.Value ?? ""));

        api.MapGet("/menu", (ClaimsPrincipal user, MenuPermissionService perms) =>
            MenuFor(user.FindFirst(ClaimTypes.Role)?.Value ?? "", perms.IsAllowed));

        api.MapGet("/pc-download", (ClaimsPrincipal user, MenuPermissionService perms, IWebHostEnvironment env) =>
            perms.IsAllowed(user, AppMenus.PcDownload)
                ? Results.Ok(ToPortalDownload(StandaloneDownloadCatalog.Read(env.ContentRootPath)))
                : Results.StatusCode(StatusCodes.Status403Forbidden));

        app.MapAdminUsers();
        app.MapAdminTools();
        app.MapToolPages();
        app.MapTeamMemo();
        app.MapProductionInput();
        app.MapBmesTables();
        app.MapQr();
        app.MapSetting();

        // Every /app route is the React app; it asks /api/portal/me and sends signed-out users to /login.
        app.MapFallbackToFile("/app/{*path:nonfile}", "app/index.html", ShellFileOptions);

        return app;
    }
}
