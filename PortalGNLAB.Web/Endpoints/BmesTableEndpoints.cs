using System.Text.Json;
using System.Threading.Channels;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// API of the React pages /app/bmes/routing-table, reason-table, worker-status and make-model-group
/// (ports of BmesRoutingTablePage / BmesReasonTablePage / WorkerStatusPage / BmesMakeModelGroupPage.razor).
/// Each group needs the same menu permission as the Blazor page. Only local SQLite is written here;
/// nothing is written to the BMES SQL Server.
/// </summary>
public static class BmesTableEndpoints
{
    private const string RequestedWithHeader = "X-Requested-With";
    private const string RequestedWithValue = "PortalGNLAB";

    public sealed record RoutingEdit(string ModelName, string ProcessCode, string ProcessName, string ProcessType);

    public sealed record ReasonEdit(string ProcessName, string NgName, string Reason);

    public sealed record WorkerRow(
        DateTime Date, string EmpNo, string Name, string DepartmentNo, string DepartmentCode, string DepartmentName,
        string WorkStatus, string DayType, string SchedStart, string SchedEnd, string CheckIn, string CheckOut);

    public sealed record WorkerExport(DateTime Start, DateTime End, List<WorkerRow> Records);

    public sealed record SubGroupDto(string Name, List<string> LineShifts, List<SubGroupDto> SubGroups);

    public sealed record MidGroupDto(string Material, List<SubGroupDto> SubGroups);

    public sealed record ModelGroupDto(string Name, string ProductGroup, List<MidGroupDto> MidGroups);

    public sealed record MaterialRow(string Maktx, string Mtype);

    private static WorkerRow ToRow(WorkerStatusService.WorkerRecord r) => new(
        r.Date, r.EmpNo, r.Name, r.DepartmentNo, r.DepartmentCode, r.DepartmentName,
        r.WorkStatus, r.DayType, r.SchedStart, r.SchedEnd, r.CheckIn, r.CheckOut);

    private static WorkerStatusService.WorkerRecord ToRecord(WorkerRow r) => new()
    {
        Date = r.Date, EmpNo = r.EmpNo ?? "", Name = r.Name ?? "", DepartmentNo = r.DepartmentNo ?? "",
        DepartmentCode = r.DepartmentCode ?? "", DepartmentName = r.DepartmentName ?? "",
        WorkStatus = r.WorkStatus ?? "", DayType = r.DayType ?? "", SchedStart = r.SchedStart ?? "",
        SchedEnd = r.SchedEnd ?? "", CheckIn = r.CheckIn ?? "", CheckOut = r.CheckOut ?? "",
    };

    private static SubGroupDto ToDto(SubGroupRecord s) => new(s.Name ?? "", s.LineShifts.ToList(), s.SubGroups.Select(ToDto).ToList());

    private static SubGroupRecord ToRecord(SubGroupDto s) => new()
    {
        Name = s.Name ?? "",
        LineShifts = (s.LineShifts ?? []).ToList(),
        SubGroups = (s.SubGroups ?? []).Select(ToRecord).ToList(),
    };

    private static int CountSubs(IEnumerable<SubGroupRecord> list) => list.Sum(s => 1 + CountSubs(s.SubGroups));

    private static object ModelGroupView(WebRepository repo) => new
    {
        groups = repo.GetModelGroups().Select(g => new ModelGroupDto(
            g.Name, g.ProductGroup,
            g.MidGroups.Select(m => new MidGroupDto(m.Material, m.SubGroups.Select(ToDto).ToList())).ToList())).ToList(),
        rows = repo.GetBmesMaktxMtypeDistinct().Select(r => new MaterialRow(r.Maktx, r.Mtype)).ToList(),
        mtypeMap = repo.GetMtypeCategoryMap(),
    };

    /// <summary>IProgress that is safe to call from any thread and keeps order (Progress&lt;T&gt; may reorder).</summary>
    private sealed class ChannelProgress(ChannelWriter<string> writer) : IProgress<string>
    {
        public void Report(string value) => writer.TryWrite(value);
    }

    /// <summary>Runs <paramref name="work"/> and answers as NDJSON: {type:"progress"} lines, then one {type:"result"} or {type:"error"}.</summary>
    private static async Task Stream(HttpContext ctx, Func<IProgress<string>, Task<object?>> work)
    {
        var channel = Channel.CreateUnbounded<string>();
        var task = Task.Run(async () =>
        {
            try { return await work(new ChannelProgress(channel.Writer)); }
            finally { channel.Writer.TryComplete(); }
        });

        ctx.Response.ContentType = "application/x-ndjson";
        async Task Write(object line)
        {
            await ctx.Response.WriteAsync(JsonSerializer.Serialize(line, JsonSerializerOptions.Web) + "\n", ctx.RequestAborted);
            await ctx.Response.Body.FlushAsync(ctx.RequestAborted);
        }

        await foreach (string message in channel.Reader.ReadAllAsync(ctx.RequestAborted))
            await Write(new { type = "progress", message });
        try { await Write(new { type = "result", data = await task }); }
        catch (Exception ex) when (ex is not OperationCanceledException) { await Write(new { type = "error", message = ex.Message }); }
    }

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

    private static object SyncResult(BmesSettingsSyncResult r) => new { succeeded = r.Succeeded, message = r.Message, rows = r.Rows };

    public static IEndpointRouteBuilder MapBmesTables(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/bmes").RequireAuthorization();

        // ── Routing Table ─────────────────────────────────────────────────────
        var routing = Guard(api.MapGroup("/routing-table"), AppMenus.BmesRoutingTable);
        routing.MapGet("", (NgRateSettingsService s) => Results.Ok(new { rows = s.GetRoutingRows() }));
        routing.MapPost("", (RoutingEdit body, NgRateSettingsService s) =>
        {
            var row = new RoutingRow
            {
                ModelName = body.ModelName ?? "", ProcessCode = body.ProcessCode ?? "",
                ProcessName = body.ProcessName ?? "", ProcessType = body.ProcessType ?? "",
            };
            return Results.Ok(new { id = s.AddRoutingRow(row) });
        });
        routing.MapPut("/{id:long}", (long id, RoutingEdit body, NgRateSettingsService s) =>
        {
            s.UpdateRoutingRow(new RoutingRow
            {
                Id = id, ModelName = body.ModelName ?? "", ProcessCode = body.ProcessCode ?? "",
                ProcessName = body.ProcessName ?? "", ProcessType = body.ProcessType ?? "",
            });
            return Results.NoContent();
        });
        routing.MapDelete("/{id:long}", (long id, NgRateSettingsService s) =>
        {
            s.DeleteRoutingRow(id);
            return Results.NoContent();
        });
        routing.MapPost("/import", (NgRateSettingsService s) =>
            Results.Ok(new { count = s.ImportRoutingFromFile(s.RoutingFilePath) }));
        routing.MapPost("/refresh", (HttpContext ctx, BmesRoutingScrapeService scrape) =>
            Stream(ctx, async progress => new { saved = await scrape.FetchAllAsync(progress) }));
        routing.MapPost("/sync", async (NgRateSettingsService s, BmesSettingsSyncService sync) =>
            Results.Ok(SyncResult(await sync.SyncRoutingRowsAsync(s.GetRoutingRows()))));

        // ── Reason Table ──────────────────────────────────────────────────────
        var reason = Guard(api.MapGroup("/reason-table"), AppMenus.BmesReasonTable);
        reason.MapGet("", (NgRateSettingsService s) => Results.Ok(new { rows = s.GetReasonRows() }));
        reason.MapPost("", (ReasonEdit body, NgRateSettingsService s) =>
            Results.Ok(new { id = s.AddReasonRow(new ReasonRow { ProcessName = body.ProcessName ?? "", NgName = body.NgName ?? "", Reason = body.Reason ?? "" }) }));
        reason.MapPut("/{id:long}", (long id, ReasonEdit body, NgRateSettingsService s) =>
        {
            s.UpdateReasonRow(new ReasonRow { Id = id, ProcessName = body.ProcessName ?? "", NgName = body.NgName ?? "", Reason = body.Reason ?? "" });
            return Results.NoContent();
        });
        reason.MapDelete("/{id:long}", (long id, NgRateSettingsService s) =>
        {
            s.DeleteReasonRow(id);
            return Results.NoContent();
        });
        reason.MapPost("/import", (NgRateSettingsService s) =>
            Results.Ok(new { count = s.ImportReasonFromFile(s.ReasonFilePath) }));
        reason.MapGet("/model-pairs", async (string model, NgRateSettingsService s) =>
        {
            var pairs = await Task.Run(() => s.GetProcessNgPairsByMaterial(model));
            return Results.Ok(new { pairs = pairs.Select(p => new { processName = p.ProcessName, ngName = p.NgName }) });
        });
        // Same two steps as the Blazor page: BMES API pairs first, then the cached daily NG data.
        reason.MapPost("/find-pairs", (HttpContext ctx, NgRateSettingsService s, BmesRoutingScrapeService scrape) =>
            Stream(ctx, async progress =>
            {
                int apiAdded = 0, cacheAdded = 0;
                string? apiFailure = null, cacheFailure = null;
                string lastMessage = "";
                try
                {
                    int r = await scrape.FetchReasonPairsAsync(new DirectProgress(m => { lastMessage = m; progress.Report(m); }));
                    if (r < 0)
                        apiFailure = string.IsNullOrWhiteSpace(lastMessage)
                            ? "BMES API step failed"
                            : $"BMES API step failed: {lastMessage.TrimEnd('.')}";
                    else apiAdded = r;
                }
                catch (Exception ex) { apiFailure = $"BMES API step failed: {ex.Message}"; }

                try { cacheAdded = await Task.Run(() => s.RefreshReasonFromDailyDbs(progress)); }
                catch (Exception ex) { cacheFailure = $"Cached NG data step failed: {ex.Message}"; }

                return (object?)new { apiAdded, cacheAdded, apiFailure, cacheFailure };
            }));
        reason.MapPost("/sync", async (NgRateSettingsService s, BmesSettingsSyncService sync) =>
            Results.Ok(SyncResult(await sync.SyncReasonRowsAsync(s.GetReasonRows()))));

        // ── Worker Status ─────────────────────────────────────────────────────
        var worker = Guard(api.MapGroup("/worker-status"), AppMenus.BmesWorkerStatus);
        worker.MapPost("/fetch", (HttpContext ctx, WorkerStatusService svc, WorkerFetch body) =>
            Stream(ctx, async progress =>
            {
                var result = await svc.FetchRangeAsync(body.Start.Date, body.End.Date, progress);
                return new
                {
                    isSuccess = result.IsSuccess,
                    errorMessage = result.ErrorMessage,
                    totalCount = result.TotalCount,
                    fetchedAt = result.FetchedAt,
                    records = result.Records.Select(ToRow).ToList(),
                };
            }));
        worker.MapPost("/export", (WorkerExport body) =>
        {
            if (body.Records is null || body.Records.Count == 0) return Results.BadRequest(new { error = "No records to export." });
            var bytes = WorkerStatusExcelExporter.Export(body.Records.Select(ToRecord).ToList(), body.Start, body.End);
            string name = body.Start.Date == body.End.Date
                ? $"WorkerStatus_{body.Start:yyyyMMdd}.xlsx"
                : $"WorkerStatus_{body.Start:yyyyMMdd}_{body.End:yyyyMMdd}.xlsx";
            return Results.File(bytes, ProductionInputEndpoints.XlsxContentType, name);
        });

        // ── Model Group ───────────────────────────────────────────────────────
        var groups = Guard(api.MapGroup("/model-groups"), AppMenus.BmesMakeModelGroup);
        groups.MapGet("", (WebRepository repo) => Results.Ok(ModelGroupView(repo)));
        groups.MapGet("/line-shifts", async (string material, NgRateSettingsService s) =>
            Results.Ok(new { lineShifts = await Task.Run(() => s.GetLineShiftsByMaterialName(material)) }));
        groups.MapPost("/pull", async (WebRepository repo, BmesSettingsSyncService sync) =>
        {
            var groupResult = await sync.PullModelGroupsFromServerAsync();
            var materialResult = await sync.PullBmesMaterialsFromServerAsync();
            return Results.Ok(new
            {
                groupResult = SyncResult(groupResult),
                materialResult = SyncResult(materialResult),
                view = ModelGroupView(repo),
            });
        });
        groups.MapPut("", async (List<ModelGroupDto> body, WebRepository repo) =>
        {
            string[] productGroups = ["SPK", "UNIT", "MODULE", "TWS", "ETC"];
            var records = (body ?? []).Select((g, i) => new ModelGroupRecord
            {
                Name = g.Name ?? "",
                ProductGroup = productGroups.Contains(g.ProductGroup) ? g.ProductGroup : "ETC",
                SortOrder = i,
                MidGroups = (g.MidGroups ?? []).Select(m => new MidGroupRecord
                {
                    Material = m.Material ?? "",
                    SubGroups = (m.SubGroups ?? []).Select(ToRecord).ToList(),
                }).ToList(),
            }).ToList();

            await Task.Run(() => repo.SaveModelGroups(records));
            return Results.Ok(new
            {
                groups = records.Count,
                midGroups = records.Sum(r => r.MidGroups.Count),
                subGroups = records.Sum(r => r.MidGroups.Sum(m => CountSubs(m.SubGroups))),
                lineShifts = records.Sum(r => r.MidGroups.Sum(m => m.LineShifts.Count)),
            });
        });

        return app;
    }

    public sealed record WorkerFetch(DateTime Start, DateTime End);

    private sealed class DirectProgress(Action<string> report) : IProgress<string>
    {
        public void Report(string value) => report(value);
    }
}
