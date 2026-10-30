using System.Globalization;
using System.Text.Json;
using System.Threading.Channels;
using PortalGNLAB.Web.Services;
using PortalGNLAB.Web.Services.BmesReports;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// API of the React report pages /app/bmes/daily-report, /app/report/bmes, /app/bmes/test5 and
/// /app/bmes/f-cost (ports of BmesDailyReportPage / BmesReportPage / BmesTest5Page / BmesFCostPage.razor).
/// The existing services do all the work unchanged; BMES SQL Server is only read (SELECT). The Razor
/// pages stay in place — BmesReportHtmlExportService still renders BmesFCostPage statically.
/// Each group needs the same menu permission as its Blazor page.
/// </summary>
public static class BmesReportEndpoints
{
    private const string RequestedWithHeader = "X-Requested-With";
    private const string RequestedWithValue = "PortalGNLAB";

    // Same setting keys as BmesReportPage.razor, so the Blazor and React pages share one selection.
    private const string SelectionSettingKey = "BmesReport:SelectionState";
    private const string ProgressBaselineKey = "BmesReport:StageBaselines";
    private const string DailyDefaultAssetVersion = "20260820-1";
    private const int MaxBomRowsPerModel = 2000;

    private static readonly JsonSerializerOptions SelectionJsonOptions = new() { PropertyNameCaseInsensitive = true };

    // ── DTOs ────────────────────────────────────────────────────────────────

    public sealed record PickerEntry(string Key, string Label, int LsCount, int Depth);

    public sealed record PickerGroup(long Id, string Name, string ProductGroup, int LsCount, List<PickerEntry> Entries);

    public sealed record Selection(List<long> GroupIds, List<string> SubEntries);

    public sealed record GenerateRequest(DateTime Start, DateTime End, List<long>? GroupIds, List<string>? SubEntries);

    public sealed record StageDto(
        string Name, string Status, string CurrentStep, double? Percent,
        int? StepCurrent, int? StepTotal, string? PhaseName, double ElapsedSeconds);

    public sealed record FCostGenerateRequest(DateTime Start, DateTime End);

    public sealed record FCostColumnDto(string Header, string Kind);

    public sealed record FCostBlockDto(string DisplayName, string ProductGroup, string Material, double[] Input, double[] Cost);

    public sealed record FCostSubGroupDto(
        string Display, string GroupName, string MidGroupName, string SubGroupKey, int MatchedMaterialCount,
        double[] InputByCol, double[] CostByCol, double[] SourceRateByCol);

    public sealed record FCostRawStatusDto(bool Exists, int SuccessCount, int FailedCount, int TotalRows, string FirstDate, string LastDate);

    public sealed record BomLoadRequest(List<string>? Models, string? Plant, DateTime WorkDate, int MaxDepth, bool FromServer);

    public sealed record CatalogSyncRequest(string? Plant, DateTime WorkDate);

    public sealed record PdmMatchRequest(List<string>? PartNames, List<string>? ModelNames);

    public sealed record CostRequest(List<string>? Codes);

    public sealed record DrawingRequest(string? Pdmno, string? Pdmvr);

    public sealed record ActivityRequest(string? Message);

    // Persisted shape of BmesReportPage.razor's selection (PascalCase on disk, as written there).
    private sealed record SelectionState(List<SelectionRoot>? RootGroups, List<SelectionSub>? SubGroups);

    private sealed record SelectionRoot(long GroupId, string GroupName);

    private sealed record SelectionSub(long GroupId, string SelectionKey, string GroupName, string Kind, string Material, string Path);

    /// <summary>IProgress that is safe to call from any thread and keeps order (Progress&lt;T&gt; may reorder).</summary>
    private sealed class ChannelProgress(ChannelWriter<object> writer, Func<string, object> line) : IProgress<string>
    {
        public void Report(string value) => writer.TryWrite(line(value));
    }

    // ── streaming helper ────────────────────────────────────────────────────

    /// <summary>
    /// Runs <paramref name="work"/> and answers as NDJSON: the lines it writes to the channel (progress,
    /// log, stages), then one {type:"result"} or {type:"error"}. When the browser goes away the work is
    /// still awaited, so the scoped services it uses are not disposed under it.
    /// </summary>
    private static async Task Stream(HttpContext ctx, Func<ChannelWriter<object>, Task<object?>> work)
    {
        var channel = Channel.CreateUnbounded<object>();
        var task = Task.Run(async () =>
        {
            try { return await work(channel.Writer); }
            finally { channel.Writer.TryComplete(); }
        });

        ctx.Response.ContentType = "application/x-ndjson";
        ctx.Response.Headers.CacheControl = "no-store";
        async Task Write(object line)
        {
            await ctx.Response.WriteAsync(JsonSerializer.Serialize(line, JsonSerializerOptions.Web) + "\n", ctx.RequestAborted);
            await ctx.Response.Body.FlushAsync(ctx.RequestAborted);
        }

        try
        {
            await foreach (object line in channel.Reader.ReadAllAsync(ctx.RequestAborted))
                await Write(line);
            await Write(new { type = "result", data = await task });
        }
        catch (OperationCanceledException) when (ctx.RequestAborted.IsCancellationRequested)
        {
            try { await task; } catch { /* nobody is listening any more */ }
        }
        catch (Exception ex)
        {
            try { await Write(new { type = "error", message = ex.Message }); }
            catch (Exception) when (ctx.RequestAborted.IsCancellationRequested) { }
        }
    }

    private static object ProgressLine(string message) => new { type = "progress", message };

    private static object LogLine(string message) => new { type = "log", message = $"[{DateTime.Now:HH:mm:ss}] {message}" };

    private static RouteGroupBuilder Guard(RouteGroupBuilder group, string menuId)
    {
        group.AddEndpointFilter(async (ctx, next) =>
        {
            var http = ctx.HttpContext;
            if (!http.RequestServices.GetRequiredService<MenuPermissionService>().IsAllowed(http.User, menuId))
                return Results.StatusCode(StatusCodes.Status403Forbidden);
            if (!HttpMethods.IsGet(http.Request.Method) && http.Request.Headers[RequestedWithHeader] != RequestedWithValue)
                return Results.Json(new { error = "Required request header is missing." }, statusCode: StatusCodes.Status400BadRequest);
            try { return await next(ctx); }
            catch (ArgumentException ex) { return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status400BadRequest); }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
            }
        });
        return group;
    }

    private static IResult BadRequest(string message) => Results.Json(new { error = message }, statusCode: StatusCodes.Status400BadRequest);

    public static IEndpointRouteBuilder MapBmesReports(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/bmes-report").RequireAuthorization();

        MapDaily(Guard(api.MapGroup("/daily"), AppMenus.BmesDailyReport));
        MapReport(Guard(api.MapGroup("/report"), AppMenus.Report));
        MapTest5(Guard(api.MapGroup("/test5"), AppMenus.BmesTest5));
        MapFCost(Guard(api.MapGroup("/f-cost"), AppMenus.BmesFCost));

        return app;
    }

    // ── DAILY REPORT ────────────────────────────────────────────────────────

    /// <summary>The page hosts the standalone viewer (ClientApp/DailyReport → wwwroot/daily-report) the
    /// same way BmesDailyReportPage.razor does; its data still comes from /api/bmes/daily-report.</summary>
    private static void MapDaily(RouteGroupBuilder daily)
    {
        daily.MapGet("/viewer", (IConfiguration configuration, IWebHostEnvironment env) =>
        {
            string version = configuration["DailyReport:ViewerAssetVersion"]?.Trim() is { Length: > 0 } configured
                ? configured
                : DailyDefaultAssetVersion;
            string webRoot = env.WebRootPath ?? Path.Combine(env.ContentRootPath, "wwwroot");
            bool assetsExist = NonEmptyFile(Path.Combine(webRoot, "daily-report", "daily-report.js"))
                && NonEmptyFile(Path.Combine(webRoot, "daily-report", "daily-report.css"));
            return new
            {
                reactEnabled = configuration.GetValue<bool>("DailyReport:ReactViewerEnabled"),
                assetsExist,
                assetVersion = version,
                scriptUrl = $"/daily-report/daily-report.js?v={Uri.EscapeDataString(version)}",
                cssUrl = $"/daily-report/daily-report.css?v={Uri.EscapeDataString(version)}",
                dataUrl = "/api/bmes/daily-report",
                refreshUrl = "/api/bmes/daily-report/refresh",
            };
        });

        daily.MapPost("/viewer-failed", (ActivityRequest body, AppActivityLogger activity) =>
        {
            activity.Log("BmesDailyReport", "React viewer mount failed (/app): " + (body.Message ?? ""));
            return Results.NoContent();
        });
    }

    private static bool NonEmptyFile(string path)
    {
        try { return new FileInfo(path) is { Exists: true, Length: > 0 }; }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    // ── REPORT (/report/bmes) ───────────────────────────────────────────────

    private static void MapReport(RouteGroupBuilder report)
    {
        report.MapGet("/setup", (WebRepository repo, NgRateSettingsService settings) =>
        {
            List<ModelGroupRecord> groups = repo.GetModelGroups();
            Selection selection = LoadSelection(repo, groups);
            return new
            {
                credentialsConfigured = settings.IsCredentialsConfigured,
                loginId = settings.LoginId,
                groups = groups.Select(ToPickerGroup).ToList(),
                productGroups = ModelGroupPickerHelpers.ProductGroupsWithRows(groups).ToList(),
                activeProductGroup = ModelGroupPickerHelpers.FirstProductGroupWithRows(groups),
                selection,
                stageNames = BmesReportHtmlExportService.StageNames,
            };
        });

        // Saved on every toggle, like the Blazor page; unknown keys are dropped.
        report.MapPut("/selection", (Selection body, WebRepository repo) =>
        {
            List<ModelGroupRecord> groups = repo.GetModelGroups();
            Selection valid = ValidSelection(groups, body);
            SaveSelection(repo, groups, valid);
            return valid;
        });

        report.MapPost("/generate", async (GenerateRequest body, HttpContext ctx, WebRepository repo,
            NgRateSettingsService settings, BmesReportHtmlExportService htmlExport) =>
        {
            DateTime start = body.Start.Date;
            DateTime end = body.End.Date;
            if (end < start) return BadRequest("End date must be on or after start date.");
            if (!settings.IsCredentialsConfigured) return BadRequest("BMES credentials are not configured.");

            List<ModelGroupRecord> groups = repo.GetModelGroups();
            Selection selection = ValidSelection(groups, new Selection(body.GroupIds ?? [], body.SubEntries ?? []));
            if (selection.GroupIds.Count == 0 && selection.SubEntries.Count == 0)
                return BadRequest("Select at least one sub group.");

            List<ModelGroupRecord> selectedGroups = ModelGroupPickerHelpers.ApplyPickerSelection(
                groups, selection.GroupIds.ToHashSet(), selection.SubEntries.ToHashSet(StringComparer.Ordinal));

            await Stream(ctx, async writer =>
            {
                int dirty = 0;
                var tracker = new ReportProgressTracker(BmesReportHtmlExportService.StageNames,
                    onChanged: () => Interlocked.Exchange(ref dirty, 1));
                tracker.ImportBaselines(LoadStageBaselines(repo));

                object StagesLine() => new
                {
                    type = "stages",
                    overall = tracker.OverallPercent,
                    stages = tracker.Stages.Select(ToStageDto).ToList(),
                };

                // Stage bars change on every progress message; send at most a few snapshots a second.
                using var tickerStop = new CancellationTokenSource();
                Task ticker = Task.Run(async () =>
                {
                    try
                    {
                        while (!tickerStop.IsCancellationRequested)
                        {
                            await Task.Delay(300, tickerStop.Token);
                            if (Interlocked.Exchange(ref dirty, 0) == 1) writer.TryWrite(StagesLine());
                        }
                    }
                    catch (OperationCanceledException) { }
                });

                void AddLog(string message) => writer.TryWrite(LogLine(message));

                AddLog($"Report start: {start:yyyy-MM-dd} - {end:yyyy-MM-dd}, selections={selection.GroupIds.Count + selection.SubEntries.Count:N0}");
                try
                {
                    string token = await htmlExport.GenerateAllTabsAsync(start, end, selectedGroups, tracker, AddLog);
                    tracker.SettleUnfinished(failed: false);
                    Dictionary<string, int> baselines = tracker.ExportBaselines();
                    if (baselines.Count > 0)
                        repo.SetSetting(ProgressBaselineKey, JsonSerializer.Serialize(baselines, SelectionJsonOptions));
                    AddLog("Report complete.");
                    return new { token };
                }
                catch (Exception ex)
                {
                    AddLog("[ERROR] " + ex);
                    tracker.SettleUnfinished(failed: true);
                    throw;
                }
                finally
                {
                    tickerStop.Cancel();
                    await ticker;
                    writer.TryWrite(StagesLine());
                }
            });
            return Results.Empty;
        });
    }

    private static StageDto ToStageDto(ReportStage s) => new(
        s.Name, s.Status.ToString(), s.CurrentStep, s.Percent, s.StepCurrent, s.StepTotal, s.PhaseName,
        Math.Round(s.Elapsed.TotalSeconds, 1));

    private static PickerGroup ToPickerGroup(ModelGroupRecord g) => new(
        g.Id, g.Name, g.ProductGroup, g.MidGroups.Sum(m => m.LineShifts.Count),
        ModelGroupPickerHelpers.EnumerateChildEntries(g)
            .Select(e => new PickerEntry(e.Key, e.Label, e.LsCount, e.Depth)).ToList());

    private static Dictionary<string, int>? LoadStageBaselines(WebRepository repo)
    {
        string? value = repo.GetSetting(ProgressBaselineKey);
        if (string.IsNullOrWhiteSpace(value)) return null;
        try { return JsonSerializer.Deserialize<Dictionary<string, int>>(value, SelectionJsonOptions); }
        catch { return null; }
    }

    private static Selection ValidSelection(IReadOnlyList<ModelGroupRecord> groups, Selection selection)
    {
        var validIds = groups.Select(g => g.Id).ToHashSet();
        var validEntries = groups
            .SelectMany(g => ModelGroupPickerHelpers.EnumerateChildEntries(g).Select(e => e.Key))
            .ToHashSet(StringComparer.Ordinal);
        return new Selection(
            (selection.GroupIds ?? []).Where(validIds.Contains).Distinct().OrderBy(id => id).ToList(),
            (selection.SubEntries ?? []).Where(k => k is not null && validEntries.Contains(k))
                .Distinct(StringComparer.Ordinal).OrderBy(k => k, StringComparer.Ordinal).ToList());
    }

    /// <summary>Port of BmesReportPage.LoadSelectionState: ids that no longer exist are recovered by group name.</summary>
    private static Selection LoadSelection(WebRepository repo, IReadOnlyList<ModelGroupRecord> groups)
    {
        var empty = new Selection([], []);
        string? value = repo.GetSetting(SelectionSettingKey);
        if (string.IsNullOrWhiteSpace(value)) return empty;

        SelectionState? state;
        try { state = JsonSerializer.Deserialize<SelectionState>(value, SelectionJsonOptions); }
        catch { return empty; }
        if (state is null) return empty;

        var groupNamesById = groups.ToDictionary(g => g.Id, g => g.Name);
        var groupIdsByName = new Dictionary<string, long>(StringComparer.Ordinal);
        foreach (var g in groups) groupIdsByName.TryAdd(g.Name, g.Id);

        var nextGroupIds = new HashSet<long>();
        foreach (var item in state.RootGroups ?? [])
        {
            if (groupNamesById.ContainsKey(item.GroupId)) nextGroupIds.Add(item.GroupId);
            else if (!string.IsNullOrWhiteSpace(item.GroupName) && groupIdsByName.TryGetValue(item.GroupName, out var recovered))
                nextGroupIds.Add(recovered);
        }

        var validSubKeys = groups
            .SelectMany(g => ModelGroupPickerHelpers.EnumerateChildEntries(g).Select(e => e.Key))
            .ToHashSet(StringComparer.Ordinal);
        var nextSubEntries = new HashSet<string>(StringComparer.Ordinal);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in state.SubGroups ?? [])
        {
            if (item is null || !seen.Add(item.SelectionKey ?? "")) continue;
            if (!string.IsNullOrWhiteSpace(item.SelectionKey) && validSubKeys.Contains(item.SelectionKey))
                nextSubEntries.Add(item.SelectionKey);
            else if (ResolveSubSelectionKey(item, groupNamesById, groupIdsByName, out var normalized)
                     && normalized is not null && validSubKeys.Contains(normalized))
                nextSubEntries.Add(normalized);
        }

        return new Selection(
            nextGroupIds.OrderBy(id => id).ToList(),
            nextSubEntries.OrderBy(k => k, StringComparer.Ordinal).ToList());
    }

    /// <summary>Port of BmesReportPage.SaveSelectionState (same JSON shape).</summary>
    private static void SaveSelection(WebRepository repo, IReadOnlyList<ModelGroupRecord> groups, Selection selection)
    {
        var groupById = groups.ToDictionary(g => g.Id, g => g.Name);
        var state = new SelectionState(
            RootGroups: selection.GroupIds
                .Where(groupById.ContainsKey)
                .Select(id => new SelectionRoot(id, groupById[id]))
                .OrderBy(x => x.GroupId)
                .ToList(),
            SubGroups: selection.SubEntries
                .Select(key => TryParseSelectionKey(key, out var id, out var kind, out var mat, out var path)
                    ? new SelectionSub(id, key, groupById.GetValueOrDefault(id) ?? string.Empty, kind, mat, path)
                    : null)
                .OfType<SelectionSub>()
                .OrderBy(x => x.SelectionKey, StringComparer.Ordinal)
                .ToList());
        repo.SetSetting(SelectionSettingKey, JsonSerializer.Serialize(state, SelectionJsonOptions));
    }

    private static bool ResolveSubSelectionKey(
        SelectionSub item,
        Dictionary<long, string> groupNamesById,
        Dictionary<string, long> groupIdsByName,
        out string? resolvedKey)
    {
        resolvedKey = null;

        if (TryParseSelectionKey(item.SelectionKey, out var parsedGroupId, out var kind, out var material, out var path))
        {
            if (groupNamesById.ContainsKey(parsedGroupId))
            {
                resolvedKey = item.SelectionKey;
                return true;
            }

            if (string.IsNullOrWhiteSpace(item.GroupName) || !groupIdsByName.TryGetValue(item.GroupName, out parsedGroupId))
                return false;

            resolvedKey = $"{parsedGroupId}|{kind}|{material}|{path}".TrimEnd('|');
            return !string.IsNullOrEmpty(resolvedKey);
        }

        if (string.IsNullOrWhiteSpace(item.GroupName) || string.IsNullOrWhiteSpace(item.Kind) || string.IsNullOrWhiteSpace(item.Material))
            return false;
        if (!groupIdsByName.TryGetValue(item.GroupName, out var groupId))
            return false;

        resolvedKey = $"{groupId}|{item.Kind}|{item.Material}|{item.Path}".TrimEnd('|');
        return !string.IsNullOrEmpty(resolvedKey);
    }

    private static bool TryParseSelectionKey(string? key, out long groupId, out string kind, out string material, out string path)
    {
        groupId = 0;
        kind = material = path = string.Empty;
        if (string.IsNullOrWhiteSpace(key)) return false;

        var parts = key.Split('|');
        if (parts.Length < 3 || !long.TryParse(parts[0], out groupId)) return false;
        kind = parts[1];
        material = parts[2];
        path = parts.Length > 3 ? parts[3] : string.Empty;
        return true;
    }

    // ── BOM & Drawing (/bmes/test5) ─────────────────────────────────────────

    private static void MapTest5(RouteGroupBuilder test5)
    {
        test5.MapGet("/state", (string? plant, WebRepository repo, BmesBomCacheService bomCache, BmesPdmScrapeService pdm) =>
        {
            var (pdmTotal, pdmFetchedAt) = pdm.GetStatus();
            return new
            {
                cachedModels = CachedModelEntries(repo, bomCache),
                catalog = Catalog(bomCache, plant ?? "3200"),
                pdmTotal,
                pdmFetchedAt,
            };
        });

        test5.MapGet("/catalog", (string? plant, BmesBomCacheService bomCache) => Catalog(bomCache, plant ?? "3200"));

        test5.MapPost("/catalog/sync", async (CatalogSyncRequest body, BmesFcostActualService fcostActual,
            BmesBomCacheService bomCache, AppPathsService paths) =>
        {
            var connection = BuildNetworkDbConnection(paths);
            if (string.IsNullOrWhiteSpace(connection.Server) || string.IsNullOrWhiteSpace(connection.UserId))
                return BadRequest("The BMES model-name synchronization connection is not configured.");

            string plant = (body.Plant ?? "").Trim();
            List<BmesBomModelCandidate> models = await fcostActual.FetchBomModelCatalogAsync(new BmesBomModelCatalogQuery
            {
                Connection = connection,
                Plant = plant,
                WorkDate = body.WorkDate.Date,
                ProductCodePrefixes = ["P-S-", "P-M-", "P-N-", "P-H-"],
                MaxRows = 20000,
            });
            bomCache.SaveModelCatalog(body.Plant ?? "", body.WorkDate.Date, models);
            return Results.Ok(Catalog(bomCache, body.Plant ?? ""));
        });

        test5.MapPost("/bom", async (BomLoadRequest body, HttpContext ctx, WebRepository repo,
            BmesFcostActualService fcostActual, BmesBomCacheService bomCache, AppPathsService paths, AppActivityLogger activity) =>
        {
            List<string> models = (body.Models ?? [])
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (models.Count == 0) return BadRequest("Select at least one model.");

            string plant = (body.Plant ?? "").Trim();
            int depth = Math.Clamp(body.MaxDepth, 1, 12);
            var connection = BuildNetworkDbConnection(paths);

            await Stream(ctx, async writer =>
            {
                var results = new List<object>();
                int cacheLoads = 0, serverLoads = 0;
                for (int i = 0; i < models.Count; i++)
                {
                    string model = models[i];
                    var cacheQuery = new BmesBomCacheQuery(model, plant, body.WorkDate.Date, depth, MaxBomRowsPerModel);
                    string cacheWarning = string.Empty;
                    writer.TryWrite(ProgressLine($"Loading BOM {i + 1}/{models.Count} — {model}"));

                    try
                    {
                        if (!body.FromServer)
                        {
                            try
                            {
                                BmesBomCacheEntry? cached = bomCache.TryLoad(cacheQuery);
                                if (cached is not null)
                                {
                                    results.Add(BomResult(model, cached.Rows, "", "", "cache",
                                        $"{cached.FetchedAt} · BOM date {cached.Query.WorkDate:yyyy-MM-dd}"));
                                    cacheLoads++;
                                    continue;
                                }
                            }
                            catch (Exception ex)
                            {
                                cacheWarning = "Cache read failed; loaded from server instead: " + ex.Message;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(connection.Server) || string.IsNullOrWhiteSpace(connection.UserId))
                        {
                            results.Add(BomResult(model, [], "No local cache and the network DB connection is not configured.", "", "", ""));
                            continue;
                        }

                        var rows = await fcostActual.FetchBomTreeAsync(new BmesBomMaterialQuery
                        {
                            Connection = connection,
                            ModelName = model,
                            Plant = plant,
                            MaxRows = MaxBomRowsPerModel,
                            MaxDepth = depth,
                            WorkDate = body.WorkDate,
                        });

                        try
                        {
                            bomCache.Save(cacheQuery, rows);
                        }
                        catch (Exception ex)
                        {
                            cacheWarning = string.IsNullOrWhiteSpace(cacheWarning)
                                ? "BOM loaded, but cache save failed: " + ex.Message
                                : cacheWarning + " " + ex.Message;
                        }

                        results.Add(BomResult(model, rows, "", cacheWarning, "server", DateTime.Now.ToString("O", CultureInfo.InvariantCulture)));
                        serverLoads++;
                    }
                    catch (Exception ex)
                    {
                        results.Add(BomResult(model, [], ex.Message, "", "", ""));
                    }
                }

                activity.Log("BOM & Drawing", $"BOM models={results.Count} cache={cacheLoads} server={serverLoads}");
                return new { results, cacheLoads, serverLoads, cachedModels = CachedModelEntries(repo, bomCache) };
            });
            return Results.Empty;
        });

        // PDM master rows for the materials on screen, keyed by name (BmesTest5Page.RefreshPdmMatches).
        test5.MapPost("/pdm/match", (PdmMatchRequest body, BmesPdmScrapeService pdm) =>
        {
            var matches = pdm.SearchByNames(body.PartNames ?? [], body.ModelNames ?? []);
            return new
            {
                rows = matches.Select(m => new { m.Pdmno, m.Pdmvr, m.PdtnoTx, m.Pdrtx }).ToList(),
            };
        });

        test5.MapPost("/pdm/fetch", async (HttpContext ctx, BmesPdmScrapeService pdm, AppActivityLogger activity) =>
        {
            await Stream(ctx, async writer =>
            {
                int saved = await pdm.FetchAllAsync(new ChannelProgress(writer, ProgressLine));
                var (total, fetchedAt) = pdm.GetStatus();
                if (saved >= 0) activity.Log("BOM & Drawing", $"PDM fetch rows={saved}");
                return new { saved, total, fetchedAt };
            });
        });

        test5.MapPost("/pdm/drawing", async (DrawingRequest body, HttpContext ctx, BmesPdmScrapeService pdm, AppActivityLogger activity) =>
        {
            string pdmno = body.Pdmno ?? "";
            await Stream(ctx, async writer =>
            {
                var result = await pdm.DownloadDrawingAsync(pdmno, body.Pdmvr ?? "0", new ChannelProgress(writer, ProgressLine));
                if (result.Success) activity.Log("BOM & Drawing", $"DWG download {pdmno} {result.FileName}");
                return new
                {
                    success = result.Success,
                    error = result.Error,
                    fileName = result.FileName,
                    base64 = result.Success ? Convert.ToBase64String(result.Bytes) : "",
                };
            });
        });

        // Unit prices are read-only lookups; a failure only leaves a warning above the tables.
        test5.MapPost("/costs", async (CostRequest body, BmesFcostActualService fcostActual, AppPathsService paths, AppActivityLogger activity) =>
        {
            var codes = (body.Codes ?? [])
                .Select(c => (c ?? "").Trim())
                .Where(c => c.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (codes.Count == 0) return Results.Ok(new { rows = Array.Empty<object>(), error = "" });

            var connection = BuildNetworkDbConnection(paths);
            if (string.IsNullOrWhiteSpace(connection.Server) || string.IsNullOrWhiteSpace(connection.UserId))
                return Results.Ok(new { rows = Array.Empty<object>(), error = "A network DB connection is required to look up unit prices." });

            try
            {
                var prices = await fcostActual.FetchMaterialCostsAsync(new BmesMaterialCostQuery
                {
                    Connection = connection,
                    MaterialCodes = codes,
                    ReferenceDate = DateTime.Today,
                    MaxRows = codes.Count,
                });
                activity.Log("BOM & Drawing", $"Material cost lookup codes={codes.Count} priced={prices.Count}");
                return Results.Ok(new
                {
                    rows = prices.Select(p => new
                    {
                        materialCode = p.MaterialCode.Trim(), p.UnitPrice, p.PriceCurrency, p.PriceSource, p.UnitPriceVnd,
                    }).ToList(),
                    error = "",
                });
            }
            catch (Exception ex)
            {
                return Results.Ok(new { rows = Array.Empty<object>(), error = "Unit price lookup failed: " + ex.Message });
            }
        });

        test5.MapPost("/activity", (ActivityRequest body, AppActivityLogger activity) =>
        {
            activity.Log("BOM & Drawing", body.Message ?? "");
            return Results.NoContent();
        });
    }

    private static object BomResult(string model, List<BmesBomMaterialCandidate> rows, string error, string warning, string source, string fetchedAt) => new
    {
        model,
        rows = rows.Select(r => new
        {
            r.ProductCode, r.ProductName, r.MaterialCode, r.MaterialName, r.BomLevel, r.BomPath, r.BomPathText, r.HasChildren,
        }).ToList(),
        error,
        warning,
        loadSource = source,
        fetchedAt,
    };

    /// <summary>Successfully cached BOM models, with the model groups they appear in as a hint (BmesTest5Page.RefreshCachedModels).</summary>
    private static List<object> CachedModelEntries(WebRepository repo, BmesBomCacheService bomCache)
    {
        var groupsByModel = repo.GetModelGroups()
            .SelectMany(g => g.MidGroups.Select(m => (Group: g.Name, Model: (m.Material ?? string.Empty).Trim())))
            .Where(x => x.Model.Length > 0)
            .GroupBy(x => x.Model, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => string.Join(", ", g.Select(x => x.Group).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n => n, StringComparer.OrdinalIgnoreCase)),
                StringComparer.OrdinalIgnoreCase);

        return bomCache.ListSuccessfulModels()
            .Select(cache =>
            {
                string hint = groupsByModel.GetValueOrDefault(cache.ModelName) ?? string.Empty;
                string cacheHint = $"{cache.RowCount:N0} BOM row(s), cached {cache.FetchedAt}";
                return new { name = cache.ModelName, groups = string.IsNullOrWhiteSpace(hint) ? cacheHint : hint + " · " + cacheHint };
            })
            .OrderBy(m => m.name, StringComparer.OrdinalIgnoreCase)
            .Cast<object>()
            .ToList();
    }

    private static object Catalog(BmesBomCacheService bomCache, string plant)
    {
        BmesBomModelCatalogEntry entry = bomCache.LoadModelCatalog(plant);
        return new
        {
            models = entry.Models.Select(m => new { m.ProductCode, m.ProductName }).ToList(),
            syncedAt = entry.SyncedAt,
            workDate = entry.WorkDate == default ? "" : entry.WorkDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        };
    }

    /// <summary>Same connection BmesTest5Page.razor builds from the Admin DB Query settings (read-only use).</summary>
    private static BmesFcostDbConnection BuildNetworkDbConnection(AppPathsService paths)
    {
        var c = paths.Current;
        return new BmesFcostDbConnection
        {
            Server = c.AdminDbQueryServer,
            Port = c.AdminDbQueryPort > 0 ? c.AdminDbQueryPort : 1430,
            Database = string.IsNullOrWhiteSpace(c.AdminDbQueryDatabase) ||
                       string.Equals(c.AdminDbQueryDatabase, "master", StringComparison.OrdinalIgnoreCase)
                ? "BMES_LIV"
                : c.AdminDbQueryDatabase,
            UserId = c.AdminDbQueryUserId,
            Password = c.AdminDbQueryPassword,
            TimeoutSeconds = Math.Clamp(Math.Max(c.AdminDbQueryTimeoutSeconds, 300), 1, 600),
            Encrypt = c.AdminDbQueryEncrypt,
            TrustServerCertificate = c.AdminDbQueryTrustServerCertificate,
        };
    }

    // ── F-Cost (/bmes/f-cost) ───────────────────────────────────────────────

    private static void MapFCost(RouteGroupBuilder fcost)
    {
        fcost.MapGet("/setup", (FCostService service, NgRateSettingsService settings) => new
        {
            credentialsConfigured = settings.IsCredentialsConfigured,
            loginId = settings.LoginId,
            rawDbPath = service.GetRawDbPath(),
            rawStatus = ToRawStatus(service.GetRawStatus()),
        });

        // Same call as BmesFCostPage.GetReportAsync: every saved model group, shared calculation service.
        fcost.MapPost("/generate", async (FCostGenerateRequest body, HttpContext ctx, WebRepository repo,
            BmesFCostReportCalculationService calculation) =>
        {
            DateTime start = body.Start.Date;
            DateTime end = body.End.Date;
            IReadOnlyList<ModelGroupRecord> groups = repo.GetModelGroups();

            await Stream(ctx, async writer =>
            {
                BmesFCostCalculationSnapshot snapshot = await calculation.CalculateAsync(
                    BmesReportRequest.Create(start, end, groups), null, new ChannelProgress(writer, ProgressLine), CancellationToken.None);
                return ToFCostView(snapshot);
            });
            return Results.Empty;
        });

        fcost.MapPost("/activity", (ActivityRequest body, AppActivityLogger activity) =>
        {
            activity.Log("F-COST", body.Message ?? "");
            return Results.NoContent();
        });
    }

    private static FCostRawStatusDto? ToRawStatus(FCostRawStatus? s) =>
        s is null ? null : new(s.Exists, s.SuccessCount, s.FailedCount, s.TotalRows, s.FirstDate, s.LastDate);

    private static object ToFCostView(BmesFCostCalculationSnapshot snapshot)
    {
        FCostReport report = snapshot.Report;
        int count = report.ColumnHeaders.Count;

        double[] Cols(FCostRow? row)
        {
            var values = new double[count];
            if (row is null) return values;
            for (int i = 0; i < count; i++) values[i] = Finite(row.GetCol(i + 1));
            return values;
        }

        FCostBlockDto Block(FCostMaterialBlock b) => new(b.DisplayName, b.ProductGroup, b.Material, Cols(b.Input), Cols(b.Cost));

        return new
        {
            queryDate = report.QueryDate,
            totalRows = report.TotalRows,
            startDate = snapshot.StartDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            endDate = snapshot.EndDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            headers = report.ColumnHeaders,
            columns = report.Columns.Select(c => new FCostColumnDto(c.Header, c.Kind.ToString())).ToList(),
            total = report.Total is null ? null : Block(report.Total),
            materials = report.Materials.Select(Block).ToList(),
            subGroups = report.SubGroups.Select(s => new FCostSubGroupDto(
                s.Display, s.GroupName, s.MidGroupName, s.SubGroupKey, s.MatchedMaterialCount,
                s.InputByCol.Select(Finite).ToArray(), s.CostByCol.Select(Finite).ToArray(), s.SourceRateByCol.Select(Finite).ToArray())).ToList(),
            unmapped = report.UnmappedMaterials.Select(m => new { m.DisplayName, m.Material }).ToList(),
            rawStatus = ToRawStatus(snapshot.RawStatus),
        };
    }

    private static double Finite(double value) => double.IsFinite(value) ? value : 0;
}
