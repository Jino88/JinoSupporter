using System.Security.Claims;
using System.Text.Json;
using System.Threading.Channels;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// Setting API for the React page /app/setting (port of Components/Pages/BmesSettingPage.razor).
/// Anyone with the Setting menu sees App Path, BMES logins, Server Tables, History Cache and the
/// material fetch; the Codex, AI provider and QR DB cards are Admin only, as on the Blazor page.
/// </summary>
public static class SettingEndpoints
{
    public sealed record CodexSave(string ApiKey, string TranslateModel, string TranslateEffort, string ReportModel, string ReportEffort);

    public sealed record AiProvidersSave(bool CodexCliEnabled, bool CodexApiEnabled);

    public sealed record QrDbSave(
        string Server, int Port, string Database, string UserId, string Password,
        int TimeoutSeconds, int DefaultMaxRows, bool Encrypt, bool TrustServerCertificate);

    public sealed record AppPathSave(string DbRootDirectory);

    public sealed record CredentialsSave(string LoginId, string Password);

    public sealed record ApiCredentialsSave(string ApiFacility, string ApiLoginId, string ApiPassword);

    public sealed record HistoryStart(DateTime From);

    private static bool IsAdmin(ClaimsPrincipal user) => user.IsInRole(AppRoles.Admin);

    private static IResult AdminOnly() => Results.Json(new { error = "Admin only." }, statusCode: StatusCodes.Status403Forbidden);

    /// <summary>Editable copy of the saved paths; every field is kept, only the ones a card owns change.</summary>
    private static AppPathsConfig CopyPaths(AppPathsService paths) =>
        JsonSerializer.Deserialize<AppPathsConfig>(JsonSerializer.Serialize(paths.Current))!;

    private static bool PathEquals(string? a, string? b) =>
        string.Equals((a ?? string.Empty).Trim(), (b ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);

    private static object PathsView(AppPathsService paths, NgRateSettingsService ngRate) => new
    {
        dbRootDirectory = paths.Current.DbRootDirectory,
        defaultDbRootDirectory = AppPathsService.DefaultDbRootDirectory,
        ngRateFolderName = AppPathsService.NgRateFolderName,
        fCostFolderName = AppPathsService.FCostFolderName,
        workerStatusFolderName = AppPathsService.WorkerStatusFolderName,
        settingsDbPath = ngRate.SettingsDbPath,
    };

    private static object HistoryView(HistoryCacheJobService job) => new
    {
        job.IsRunning,
        job.Status,
        log = job.Log.TakeLast(60).ToList(),
        lastHistoryDay = HistoryCachePolicy.LastHistoryDay(DateTime.Today).ToString("yyyy-MM-dd"),
        liveWindowDays = HistoryCachePolicy.LiveWindowDays,
    };

    private static object MaterialView(WebRepository repo) => new
    {
        latest = repo.GetLatestBmesMaterial(),
        count = repo.GetBmesMaterialCount(),
    };

    private static object QrDbView(AppPathsConfig c) => new
    {
        server = c.QrBakoDataServer,
        port = c.QrBakoDataPort,
        database = c.QrBakoDataDatabase,
        userId = c.QrBakoDataUserId,
        password = c.QrBakoDataPassword,
        timeoutSeconds = c.QrBakoDataTimeoutSeconds,
        defaultMaxRows = c.QrBakoDataDefaultMaxRows,
        encrypt = c.QrBakoDataEncrypt,
        trustServerCertificate = c.QrBakoDataTrustServerCertificate,
    };

    private static object CodexView(WebRepository repo) => new
    {
        apiKey = repo.GetSetting("OpenAI:ApiKey") ?? repo.GetSetting("Codex:ApiKey") ?? string.Empty,
        translateModel = repo.GetSetting("OpenAI:TranslateModel") ?? repo.GetSetting("Codex:TranslateModel") ?? CodexApiService.DefaultTranslateModel,
        translateEffort = repo.GetSetting("OpenAI:TranslateEffort") is { Length: > 0 } e ? e : CodexApiService.DefaultReasoningEffort,
        reportModel = repo.GetSetting(ReviewReportService.ModelSetting) is { Length: > 0 } m ? m : ReviewReportService.DefaultModel,
        reportEffort = repo.GetSetting(ReviewReportService.EffortSetting) is { Length: > 0 } r ? r : ReviewReportService.DefaultEffort,
        reasoningEfforts = CodexApiService.ReasoningEfforts,
        lunaModel = CodexApiService.LunaModel,
        defaultReportModel = ReviewReportService.DefaultModel,
        defaultReportEffort = ReviewReportService.DefaultEffort,
    };

    private static object QrModelsView(QrModelCodeService codes) => new
    {
        codes = codes.Codes,
        path = codes.SettingsPath,
    };

    public static IEndpointRouteBuilder MapSetting(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/portal/setting").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            if (!ctx.HttpContext.RequestServices.GetRequiredService<MenuPermissionService>()
                .IsAllowed(ctx.HttpContext.User, AppMenus.BmesSetting)) return Results.Forbid();
            if (!HttpMethods.IsGet(ctx.HttpContext.Request.Method)
                && ctx.HttpContext.Request.Headers["X-Requested-With"] != "PortalGNLAB")
                return Results.BadRequest(new { error = "Required request header is missing." });
            return await next(ctx);
        });

        api.MapGet("", (ClaimsPrincipal user, AppPathsService paths, NgRateSettingsService ngRate, WebRepository repo,
            HistoryCacheJobService history, AiProviderSettingsService ai, QrModelCodeService qrCodes) =>
        {
            NgRateSettingsSnapshot snap = ngRate.GetSnapshot();
            bool admin = IsAdmin(user);
            return Results.Ok(new
            {
                isAdmin = admin,
                paths = PathsView(paths, ngRate),
                credentials = new { snap.LoginId, snap.Password, snap.ApiFacility, snap.ApiLoginId, snap.ApiPassword },
                history = HistoryView(history),
                material = MaterialView(repo),
                codex = admin ? CodexView(repo) : null,
                aiProviders = admin ? ai.GetSnapshot() : null,
                qrDb = admin ? QrDbView(paths.Current) : null,
                qrModels = admin ? QrModelsView(qrCodes) : null,
            });
        });

        // ── Admin cards ───────────────────────────────────────────────────────────────

        api.MapGet("/codex-models", async (ClaimsPrincipal user, CodexApiService codex, CancellationToken ct) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            try { return Results.Ok(new { models = await codex.ListModelsAsync(ct), error = "" }); }
            catch (Exception ex) { return Results.Ok(new { models = Array.Empty<string>(), error = $"Model list unavailable: {ex.Message}" }); }
        });

        api.MapPost("/codex", (CodexSave body, ClaimsPrincipal user, WebRepository repo, AiProviderSettingsService ai) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            string apiKey = (body.ApiKey ?? string.Empty).Trim();
            repo.SetSetting("OpenAI:ApiKey", apiKey);
            repo.SetSetting("OpenAI:TranslateModel",
                string.IsNullOrWhiteSpace(body.TranslateModel) ? CodexApiService.DefaultTranslateModel : body.TranslateModel.Trim());
            repo.SetSetting("OpenAI:TranslateEffort", body.TranslateEffort ?? CodexApiService.DefaultReasoningEffort);
            repo.SetSetting(ReviewReportService.ModelSetting,
                string.IsNullOrWhiteSpace(body.ReportModel) ? ReviewReportService.DefaultModel : body.ReportModel.Trim());
            repo.SetSetting(ReviewReportService.EffortSetting, body.ReportEffort ?? ReviewReportService.DefaultEffort);
            if (apiKey.Length > 0)
            {
                var providers = ai.GetSnapshot();
                providers.CodexApiEnabled = true;
                ai.Save(providers);
            }
            return Results.Ok(new { codex = CodexView(repo), aiProviders = ai.GetSnapshot() });
        });

        api.MapPost("/ai-providers", (AiProvidersSave body, ClaimsPrincipal user, AiProviderSettingsService ai) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            // Only the two Codex switches are on this page; the Claude ones keep their saved value.
            var providers = ai.GetSnapshot();
            providers.CodexCliEnabled = body.CodexCliEnabled;
            providers.CodexApiEnabled = body.CodexApiEnabled;
            ai.Save(providers);
            return Results.Ok(ai.GetSnapshot());
        });

        api.MapPost("/qr-db", (QrDbSave body, ClaimsPrincipal user, AppPathsService paths) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            paths.Save(ApplyQrDb(CopyPaths(paths), body));
            return Results.Ok(new { qrDb = QrDbView(paths.Current), message = "QR BAKO DB connection saved." });
        });

        api.MapPost("/qr-db/test", async (QrDbSave body, ClaimsPrincipal user, AppPathsService paths, QrTestDataService qrData, CancellationToken ct) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            paths.Save(ApplyQrDb(CopyPaths(paths), body));
            try
            {
                // One connection serves all three QR tables, so probe each of them —
                // otherwise a permission gap on ALTD/GSTD only shows up on the pages.
                foreach (QrTestTableDef table in QrTestTables.All)
                    await qrData.TestConnectionAsync(table, ct);
                return Results.Ok(new { qrDb = QrDbView(paths.Current), message = "Connected to TCDB (dbo.BKTD, dbo.ALTD, dbo.GSTD).", error = false });
            }
            catch (Exception ex)
            {
                return Results.Ok(new { qrDb = QrDbView(paths.Current), message = ex.Message, error = true });
            }
        });

        api.MapPost("/qr-model-codes", (List<QrModelCode> body, ClaimsPrincipal user, QrModelCodeService codes) =>
        {
            if (!IsAdmin(user)) return AdminOnly();
            // Blank codes and duplicates are dropped by the service, so the saved list is sent back
            // rather than the typed one — otherwise the page would keep a row the QR pages ignore.
            codes.Save(body);
            return Results.Ok(QrModelsView(codes));
        });

        // ── Every Setting user ────────────────────────────────────────────────────────

        api.MapPost("/app-path", (AppPathSave body, AppPathsService paths, NgRateSettingsService ngRate) =>
        {
            var (ok, message, restart) = SaveAppPath(body.DbRootDirectory, paths, ngRate);
            return ok
                ? Results.Ok(new { paths = PathsView(paths, ngRate), message = $"Saved. Server table DB: {ngRate.SettingsDbPath}", requiresRestart = restart })
                : Results.BadRequest(new { error = message });
        });

        api.MapPost("/credentials", (CredentialsSave body, NgRateSettingsService ngRate) =>
        {
            var snap = ngRate.GetSnapshot();
            snap.LoginId = body.LoginId ?? string.Empty;
            snap.Password = body.Password ?? string.Empty;
            ngRate.ApplySnapshot(snap);
            return Results.Ok(new { message = "Credentials saved." });
        });

        api.MapPost("/api-credentials", (ApiCredentialsSave body, NgRateSettingsService ngRate) =>
        {
            var snap = ngRate.GetSnapshot();
            snap.ApiFacility = body.ApiFacility ?? string.Empty;
            snap.ApiLoginId = body.ApiLoginId ?? string.Empty;
            snap.ApiPassword = body.ApiPassword ?? string.Empty;
            ngRate.ApplySnapshot(snap);
            return Results.Ok(new { message = "BMES API login saved." });
        });

        // Saves the working folder first (as the Blazor page did), then pulls into its settings DB.
        api.MapPost("/server-tables/{table}", async (string table, AppPathSave body, AppPathsService paths,
            NgRateSettingsService ngRate, BmesSettingsSyncService sync) =>
        {
            Func<Task<BmesSettingsSyncResult>>? pull = table switch
            {
                "routing" => sync.PullRoutingRowsFromServerAsync,
                "reason" => sync.PullReasonRowsFromServerAsync,
                "all" => sync.PullAllRowsFromServerAsync,
                _ => null,
            };
            if (pull is null) return Results.NotFound();

            var (ok, _, restart) = SaveAppPath(body.DbRootDirectory, paths, ngRate);
            if (!ok) return Results.Ok(new { succeeded = false, message = "Working folder save failed.", paths = PathsView(paths, ngRate), requiresRestart = false });
            try
            {
                var result = await pull();
                return Results.Ok(new
                {
                    succeeded = result.Succeeded,
                    message = result.Succeeded ? $"{result.Message} {result.Rows:N0} row(s) saved." : result.Message,
                    paths = PathsView(paths, ngRate),
                    requiresRestart = restart,
                });
            }
            catch (Exception ex)
            {
                return Results.Ok(new { succeeded = false, message = ex.Message, paths = PathsView(paths, ngRate), requiresRestart = restart });
            }
        });

        api.MapGet("/history", (HistoryCacheJobService history) => Results.Ok(HistoryView(history)));

        api.MapPost("/history/start", (HistoryStart body, HistoryCacheJobService history) =>
            history.Start(body.From)
                ? Results.Ok(HistoryView(history))
                : Results.Conflict(new { error = "A history cache run is already in progress." }));

        api.MapPost("/history/cancel", (HistoryCacheJobService history) =>
        {
            history.Cancel();
            return Results.Ok(HistoryView(history));
        });

        // Streams the fetch log as NDJSON ({log}) and ends with {done, saved, material}.
        api.MapPost("/materials/fetch", async (HttpContext ctx, BmesMaterialService materials, WebRepository repo) =>
        {
            ctx.Response.ContentType = "application/x-ndjson";
            var lines = Channel.CreateUnbounded<string>();
            var progress = new SyncProgress(msg => lines.Writer.TryWrite(JsonSerializer.Serialize(new { log = msg }, JsonOptions)));
            var work = Task.Run(async () =>
            {
                try
                {
                    int saved = await materials.FetchAllAsync(progress);
                    lines.Writer.TryWrite(JsonSerializer.Serialize(new { done = true, saved, material = MaterialView(repo) }, JsonOptions));
                }
                catch (Exception ex)
                {
                    lines.Writer.TryWrite(JsonSerializer.Serialize(new { log = ex.ToString() }, JsonOptions));
                    lines.Writer.TryWrite(JsonSerializer.Serialize(new { done = true, saved = 0, error = $"Unexpected error: {ex.Message}", material = MaterialView(repo) }, JsonOptions));
                }
                finally
                {
                    lines.Writer.TryComplete();
                }
            });
            await foreach (string line in lines.Reader.ReadAllAsync())
            {
                await ctx.Response.WriteAsync(line + "\n");
                await ctx.Response.Body.FlushAsync();
            }
            await work;
        });

        return app;
    }

    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    /// <summary>Reports on the calling thread, unlike <see cref="Progress{T}"/>, so lines keep their order.</summary>
    private sealed class SyncProgress(Action<string> report) : IProgress<string>
    {
        public void Report(string value) => report(value);
    }

    private static AppPathsConfig ApplyQrDb(AppPathsConfig c, QrDbSave body)
    {
        c.QrBakoDataServer = (body.Server ?? string.Empty).Trim();
        c.QrBakoDataPort = Math.Clamp(body.Port, 1, 65535);
        c.QrBakoDataDatabase = (body.Database ?? string.Empty).Trim();
        c.QrBakoDataUserId = (body.UserId ?? string.Empty).Trim();
        c.QrBakoDataPassword = body.Password ?? string.Empty;
        c.QrBakoDataTimeoutSeconds = Math.Clamp(body.TimeoutSeconds, 1, 300);
        c.QrBakoDataDefaultMaxRows = Math.Clamp(body.DefaultMaxRows, 1, 20000);
        c.QrBakoDataEncrypt = body.Encrypt;
        c.QrBakoDataTrustServerCertificate = body.TrustServerCertificate;
        return c;
    }

    /// <summary>Saves the working folder; every DB folder follows it. Restart is needed when the
    /// root, the main DB or the settings DB folder moved.</summary>
    private static (bool Ok, string Message, bool RequiresRestart) SaveAppPath(
        string? dbRootDirectory, AppPathsService paths, NgRateSettingsService ngRate)
    {
        string requested = string.IsNullOrWhiteSpace(dbRootDirectory)
            ? AppPathsService.DefaultDbRootDirectory
            : dbRootDirectory.Trim();
        if (AppPathsService.DirectoryPathConflictsWithFile(requested))
            return (false, $"Working Folder points to an existing file, not a folder: {requested}", false);

        try
        {
            AppPathsConfig before = paths.Current;
            string origRoot = before.DbRootDirectory, origMain = before.MainDbPath, origSettings = before.NgRateSettingsDbDirectory;
            AppPathsConfig form = CopyPaths(paths);
            form.DbRootDirectory = requested;
            form = AppPathsService.ApplyUnifiedDbDirectory(form);
            paths.Save(form);
            ngRate.EnsureDatabase();
            bool restart = !PathEquals(form.DbRootDirectory, origRoot)
                || !PathEquals(form.MainDbPath, origMain)
                || !PathEquals(form.NgRateSettingsDbDirectory, origSettings);
            return (true, string.Empty, restart);
        }
        catch (Exception ex)
        {
            return (false, ex.Message, false);
        }
    }
}
