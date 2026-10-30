using System.Collections.Concurrent;
using System.Globalization;
using System.Text.Json;
using System.Threading.Channels;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// API of the React NG Rate pages /app/ng-rate (By Model), /app/ng-rate-all (By Section) and
/// /app/ng-rate-by-group (By Group): ports of NgRatePage / NgRateAllPage / NgRateByGroupPage.razor.
/// The existing services fetch (BMES SELECT only) and aggregate; this file only picks the same
/// mappings the Blazor pages built, shapes the reports as JSON and keeps the last reports for the
/// Excel export. Every route needs the NG Rate menu permission.
/// </summary>
public static class NgRateEndpoints
{
    private const string AreaByModel = "NG Rate By Model";
    private const string AreaBySection = "NG Rate By Section";
    private const string AreaByGroup = "NG Rate By Group";

    public sealed record ChildEntryDto(string Key, string Label, int LsCount, int Depth, string Level);

    public sealed record PickerGroupDto(
        long Id, string Name, string ProductGroup, int MidCount, int LineShiftCount,
        List<ChildEntryDto> Children, List<ChildEntryDto> MidChildren);

    public sealed record SetupView(bool CredentialsConfigured, string LoginId, List<string> ProductGroups, List<PickerGroupDto> Groups);

    /// <summary>Mode "by-model" (whole groups as full drill-down trees) or "by-section" (flat, one section per selection).</summary>
    public sealed record ReportRequest(string? Mode, string? Start, string? End, List<long>? GroupIds, List<string>? SubEntries);

    public sealed record ByGroupRequest(string? Start, string? End, List<string>? GroupNames);

    public sealed record PeriodDto(string Key, string Header);

    public sealed record SummaryRowDto(string ProcessType, bool IsTotal, Dictionary<string, double> Ppm);

    public sealed record GroupRowDto(string GroupName, Dictionary<string, double> Ppm);

    public sealed record ProcessRowDto(int Rank, string ProcessType, string ProcessName, Dictionary<string, double> Ppm, List<GroupRowDto> Groups);

    public sealed record NgRowDto(
        int Rank, string ProcessType, string ProcessName, string NgName, Dictionary<string, double> Ppm, List<GroupRowDto> Groups);

    public sealed record ReasonRowDto(
        string Reason, int No, bool IsTotal, string ProcessType, string ProcessName, string NgName,
        Dictionary<string, double> Ppm, List<GroupRowDto> Groups);

    public sealed record ReportDto(
        List<PeriodDto> DateCols,
        List<PeriodDto> WeekCols,
        List<PeriodDto> MonthCols,
        List<SummaryRowDto> Summary,
        List<ProcessRowDto> Top10Process,
        List<NgRowDto> Top10Ng,
        List<ReasonRowDto> ReasonRows,
        List<string> GroupNames,
        Dictionary<string, List<SummaryRowDto>> GroupSummary,
        string DbPath,
        int TotalRows,
        string GeneratedAt,
        List<string> ModelFilter);

    public sealed record LabeledReportDto(string Label, ReportDto Report);

    /// <summary>One Sub / LineShift row under a model (HierReportBuilder.BuildSubRows), PPM for every period key.</summary>
    public sealed record HierSubRowDto(
        string RowId, string? ParentId, string Display, int Depth, bool HasChildren, bool IsLineShift, Dictionary<string, double> Ppm);

    /// <summary>Model row: PPM = average of its Sub1 PPMs (HierPpmLookup.MidAvgOfSubs).</summary>
    public sealed record HierMidDto(string Material, int LineShiftCount, Dictionary<string, double> Ppm, List<HierSubRowDto> Rows);

    /// <summary>Group row: PPM = average of its model PPMs (HierPpmLookup.GroupAvgOfMids).</summary>
    public sealed record HierGroupDto(string Name, Dictionary<string, double> Ppm, List<HierMidDto> Mids);

    public sealed record ReportResult(string DbPath, List<LabeledReportDto> Reports, List<HierGroupDto> Hier, string? ExportId);

    public sealed record ProcessTypeDto(string ProcessType, Dictionary<string, double> Ppm);

    public sealed record Top10RowDto(
        string Material, string ProcessType, string ProcessName, string NgName,
        Dictionary<string, double> DatePpm, Dictionary<string, double> WeekPpm, Dictionary<string, double> MonthPpm);

    public sealed record ByGroupGroupDto(
        string Name, Dictionary<string, double> Total, List<ProcessTypeDto> ProcessTypes, List<HierMidDto> Mids, List<Top10RowDto> Top10);

    public sealed record ByGroupResult(
        string DbPath,
        List<PeriodDto> DateCols,
        List<PeriodDto> WeekCols,
        List<PeriodDto> MonthCols,
        int TotalRows,
        List<ByGroupGroupDto> Groups,
        List<PeriodDto> Top10DateCols,
        List<PeriodDto> Top10WeekCols,
        List<PeriodDto> Top10MonthCols,
        string SortRefKey,
        string SortRefBucket,
        string SortRefLabel,
        string? ExportId);

    /// <summary>Expected stop of a report run; the message is shown as is (panel = also in the setup panel).</summary>
    private sealed class ReportFailure(string message, bool panel = false) : Exception(message)
    {
        public bool Panel { get; } = panel;
    }

    // ── Excel export: the reports of the last runs, so the export needs no second fetch ──────────
    private sealed record ExportEntry(
        string User,
        DateTime Created,
        string Area,
        string FilePrefix,
        List<(string Label, NgRateReportService.NgRateReport Report)> Reports,
        List<(string Label, NgRateReportService.LineShiftNgReport Report)>? Details);

    private const int MaxExports = 8;
    private static readonly TimeSpan ExportLifetime = TimeSpan.FromHours(2);
    private static readonly ConcurrentDictionary<string, ExportEntry> Exports = new(StringComparer.Ordinal);

    private static string Remember(ExportEntry entry)
    {
        DateTime oldest = DateTime.Now - ExportLifetime;
        foreach (var kv in Exports.Where(kv => kv.Value.Created < oldest).ToList())
            Exports.TryRemove(kv.Key, out _);
        foreach (var kv in Exports.OrderByDescending(kv => kv.Value.Created).Skip(MaxExports - 1).ToList())
            Exports.TryRemove(kv.Key, out _);

        string id = Guid.NewGuid().ToString("N");
        Exports[id] = entry;
        return id;
    }

    private static string UserOf(HttpContext ctx) => ctx.User.Identity?.Name ?? string.Empty;

    // ── Mapping helpers ─────────────────────────────────────────────────────────────────────────
    public static bool TryParseDay(string? text, out DateTime date) =>
        DateTime.TryParseExact(text?.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);

    private static List<PeriodDto> Cols(IEnumerable<NgRateReportService.PeriodColumn> cols) =>
        cols.Select(c => new PeriodDto(c.Key, c.Header)).ToList();

    private static IReadOnlyList<string> AllKeys(NgRateReportService.NgRateReport report) =>
        report.DateCols.Concat(report.WeekCols).Concat(report.MonthCols).Select(c => c.Key).ToList();

    /// <summary>PPM per period key; zeros (and non-finite values, which JSON cannot carry) are left out —
    /// the page shows a missing key as "-".</summary>
    private static Dictionary<string, double> Ppm(IEnumerable<string> keys, Func<string, double> read)
    {
        var ppm = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach (string key in keys)
        {
            double v = read(key);
            if (v > 0 && double.IsFinite(v)) ppm[key] = v;
        }
        return ppm;
    }

    private static Dictionary<string, double> Ppm(Dictionary<string, double> source) =>
        Ppm(source.Keys, k => source[k]);

    private static SummaryRowDto ToDto(NgRateReportService.SummaryPivotRow r) => new(r.ProcessType, r.IsTotal, Ppm(r.Ppm));

    private static List<GroupRowDto> ToDto(List<NgRateReportService.GroupPivotRow> rows) =>
        rows.Select(g => new GroupRowDto(g.GroupName, Ppm(g.Ppm))).ToList();

    public static ReportDto ToDto(NgRateReportService.NgRateReport r) => new(
        Cols(r.DateCols),
        Cols(r.WeekCols),
        Cols(r.MonthCols),
        r.Summary.Select(ToDto).ToList(),
        r.Top10Process.Select(p => new ProcessRowDto(p.Rank, p.ProcessType, p.ProcessName, Ppm(p.Ppm), ToDto(p.Groups))).ToList(),
        r.Top10Ng.Select(n => new NgRowDto(n.Rank, n.ProcessType, n.ProcessName, n.NgName, Ppm(n.Ppm), ToDto(n.Groups))).ToList(),
        r.ReasonRows.Select(x => new ReasonRowDto(x.Reason, x.No, x.IsTotal, x.ProcessType, x.ProcessName, x.NgName, Ppm(x.Ppm), ToDto(x.Groups))).ToList(),
        r.GroupNames.ToList(),
        r.GroupSummary.ToDictionary(kv => kv.Key, kv => kv.Value.Select(ToDto).ToList(), StringComparer.Ordinal),
        r.DbPath,
        r.TotalRows,
        r.GeneratedAt.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
        r.ModelFilter.ToList());

    private static HierMidDto ToHierMid(string groupName, MidGroupRecord mid, HierPpmLookup lookup, IReadOnlyList<string> keys) => new(
        mid.Material,
        mid.LineShifts.Count,
        Ppm(keys, k => lookup.MidAvgOfSubs(groupName, mid, k)),
        HierReportBuilder.BuildSubRows(groupName, mid, lookup)
            .Select(r => new HierSubRowDto(r.RowId, r.ParentId, r.Display, r.Depth, r.HasChildren, r.IsLineShift, Ppm(keys, r.PpmFor)))
            .ToList());

    /// <summary>Models without LineShifts or Material never show in the hierarchy tables.</summary>
    private static List<HierMidDto> ToHierMids(ModelGroupRecord g, HierPpmLookup lookup, IReadOnlyList<string> keys) =>
        g.MidGroups
            .Where(m => m.LineShifts.Count > 0 && !string.IsNullOrEmpty(m.Material))
            .Select(m => ToHierMid(g.Name, m, lookup, keys))
            .ToList();

    public static List<HierGroupDto> ToHier(HierReports hier)
    {
        if (!hier.HasData || hier.ByGroup is null) return [];
        var lookup = HierPpmLookup.From(hier);
        var keys = AllKeys(hier.ByGroup);
        return hier.Groups
            .Select(g => new HierGroupDto(g.Name, Ppm(keys, k => lookup.GroupAvgOfMids(g, k)), ToHierMids(g, lookup, keys)))
            .ToList();
    }

    private static ChildEntryDto ToDto(ModelGroupChildEntry e) => new(
        e.Key, e.Label, e.LsCount, e.Depth,
        e.Key.Contains("|mid|", StringComparison.Ordinal) ? "MODEL" : $"SUB {Math.Max(1, e.Depth - 1)}");

    private static SetupView Setup(WebRepository repo, NgRateSettingsService settings)
    {
        var groups = repo.GetModelGroups();
        return new SetupView(
            settings.IsCredentialsConfigured,
            settings.LoginId,
            ModelGroupPickerHelpers.ProductGroupsWithRows(groups).ToList(),
            groups.Select(g => new PickerGroupDto(
                g.Id, g.Name, g.ProductGroup, g.MidGroups.Count, g.MidGroups.Sum(m => m.LineShifts.Count),
                ModelGroupPickerHelpers.EnumerateChildEntries(g).Select(ToDto).ToList(),
                ModelGroupPickerHelpers.EnumerateMidEntries(g).Select(ToDto).ToList())).ToList());
    }

    /// <summary>Whole groups + the parents of picked children; every group when nothing is picked.</summary>
    private static List<ModelGroupRecord> ResolveHierGroups(List<ModelGroupRecord> groups, HashSet<long> groupIds, IEnumerable<string> subEntries)
    {
        var set = new HashSet<long>(groupIds);
        foreach (string key in subEntries)
        {
            string[] parts = key.Split('|');
            if (parts.Length >= 1 && long.TryParse(parts[0], out long gid)) set.Add(gid);
        }
        if (set.Count == 0)
            foreach (var g in groups) set.Add(g.Id);
        return groups.Where(g => set.Contains(g.Id)).ToList();
    }

    // ── By Group: sort reference of the Top-10 card ───────────────────────────────────────────
    private const int Top10DateCap = 10;
    private const int Top10WeekCap = 4;
    private const int Top10MonthCap = 3;

    /// <summary>Previous week (Week[1]) → Week[0] → Date[0], as NgRateByGroupPage.SortRefKey.</summary>
    public static (string Key, string Bucket) SortRefKey(NgRateReportService.NgRateReport report)
    {
        if (report.WeekCols.Count >= 2) return (report.WeekCols[1].Key, "week");
        if (report.WeekCols.Count == 1) return (report.WeekCols[0].Key, "week");
        if (report.DateCols.Count > 0) return (report.DateCols[0].Key, "date");
        return ("", "");
    }

    /// <summary>Display-only label of the sort reference: week -1, date -2 days (NgRateByGroupPage.ShiftedRefLabel).</summary>
    public static string ShiftedRefLabel(string key, string bucket)
    {
        if (string.IsNullOrEmpty(key)) return "";
        if (bucket == "week" && key.Length >= 6
            && int.TryParse(key[..4], out int yy) && int.TryParse(key[4..], out int ww))
        {
            ww -= 1;
            if (ww < 1) { yy -= 1; ww = ISOWeek.GetWeeksInYear(yy); }
            return $"{yy}{ww:D2}";
        }
        if (bucket == "date" && DateTime.TryParse(key, out DateTime dt))
            return dt.AddDays(-2).ToString("yyyy-MM-dd");
        return key;
    }

    /// <summary>Top-10 table headers: date -2 days, week -1 week; the keys stay the real columns.</summary>
    public static string ShiftedDateHeader(NgRateReportService.PeriodColumn col) =>
        DateTime.TryParse(col.Key, out DateTime dt) ? dt.AddDays(-2).ToString("MM/dd") : col.Header;

    public static string ShiftedWeekHeader(NgRateReportService.PeriodColumn col)
    {
        if (col.Key.Length >= 6
            && int.TryParse(col.Key[..4], out int yy)
            && int.TryParse(col.Key[4..], out int ww))
        {
            ww -= 1;
            if (ww < 1) { yy -= 1; ww = ISOWeek.GetWeeksInYear(yy); }
            return "W" + ww;
        }
        return col.Header;
    }

    private static ByGroupResult ToByGroup(
        NgRateGroupReportBundle bundle, List<ModelGroupRecord> groups, List<string> names, string dbPath, string? exportId)
    {
        var report = bundle.GroupReport;
        var keys = AllKeys(report);
        var lookup = bundle.ToLookup();
        var (refKey, refBucket) = SortRefKey(report);

        double MidPt(string groupName, string material, string colKey, string processType)
        {
            var rows = bundle.MidReport.GroupSummary.GetValueOrDefault(NgRateModeSupport.ModelKey(groupName, material));
            var row = rows?.FirstOrDefault(r => !r.IsTotal && r.ProcessType == processType);
            return row?.Ppm.GetValueOrDefault(colKey) ?? 0;
        }

        // Mean of the member models' PPM for one process type (values > 0 only).
        double MeanOfMembers(ModelGroupRecord g, string colKey, string processType)
        {
            double sum = 0;
            int cnt = 0;
            foreach (var mid in g.MidGroups)
            {
                double v = MidPt(g.Name, mid.Material, colKey, processType);
                if (v <= 0) continue;
                sum += v;
                cnt++;
            }
            return cnt == 0 ? 0 : sum / cnt;
        }

        double SortValue(NgRateReportService.LineShiftNgDetail d) => refBucket switch
        {
            "week" => d.WeekPpm.GetValueOrDefault(refKey),
            "date" => d.DatePpm.GetValueOrDefault(refKey),
            _ => 0,
        };

        var top10Date = report.DateCols.Take(Top10DateCap).ToList();
        var top10Week = report.WeekCols.Take(Top10WeekCap).ToList();
        var top10Month = report.MonthCols.Take(Top10MonthCap).ToList();
        Dictionary<string, double> Only(Dictionary<string, double> ppm, List<NgRateReportService.PeriodColumn> cols) =>
            Ppm(cols.Select(c => c.Key), k => ppm.GetValueOrDefault(k));

        var result = new List<ByGroupGroupDto>();
        foreach (string name in names)
        {
            var g = groups.FirstOrDefault(x => x.Name == name);
            if (g is null) continue;
            bool hasMids = g.MidGroups.Count > 0;

            var ptRows = report.GroupSummary.GetValueOrDefault(name) ?? [];
            var processTypes = ptRows
                .Where(r => !r.IsTotal)
                .OrderBy(r => NgRateModeSupport.ProcessTypeRank(r.ProcessType))
                .Select(r => new ProcessTypeDto(r.ProcessType, hasMids ? Ppm(keys, k => MeanOfMembers(g, k, r.ProcessType)) : new Dictionary<string, double>()))
                .ToList();

            string prefix = name + "::";
            var top = bundle.MidNgReport.Details
                .Where(d => d.LineShift.StartsWith(prefix, StringComparison.Ordinal))
                .Select(d => (Detail: d, Ref: SortValue(d)))
                .Where(x => x.Ref > 0)
                .OrderByDescending(x => x.Ref)
                .Take(10)
                .Select(x => new Top10RowDto(
                    x.Detail.LineShift[prefix.Length..], x.Detail.ProcessType, x.Detail.ProcessName, x.Detail.NgName,
                    Only(x.Detail.DatePpm, top10Date), Only(x.Detail.WeekPpm, top10Week), Only(x.Detail.MonthPpm, top10Month)))
                .ToList();

            result.Add(new ByGroupGroupDto(
                name,
                hasMids ? Ppm(keys, k => lookup.GroupAvgOfMids(g, k)) : new Dictionary<string, double>(),
                processTypes,
                ToHierMids(g, lookup, keys),
                top));
        }

        return new ByGroupResult(
            dbPath,
            Cols(report.DateCols),
            Cols(report.WeekCols),
            Cols(report.MonthCols),
            report.TotalRows,
            result,
            top10Date.Select(c => new PeriodDto(c.Key, ShiftedDateHeader(c))).ToList(),
            top10Week.Select(c => new PeriodDto(c.Key, ShiftedWeekHeader(c))).ToList(),
            Cols(top10Month),
            refKey,
            refBucket,
            ShiftedRefLabel(refKey, refBucket),
            exportId);
    }

    // ── Streaming ───────────────────────────────────────────────────────────────────────────────
    /// <summary>IProgress that is safe to call from any thread and keeps order (Progress&lt;T&gt; may reorder).</summary>
    private sealed class ChannelProgress(ChannelWriter<string> writer) : IProgress<string>
    {
        public void Report(string value) => writer.TryWrite(value);
    }

    /// <summary>Runs <paramref name="work"/> and answers as NDJSON: {type:"progress"} lines, then one {type:"result"} or {type:"error"}.</summary>
    private static async Task Stream(HttpContext ctx, Func<IProgress<string>, Task<object?>> work)
    {
        var channel = Channel.CreateUnbounded<string>();
        var progress = new ChannelProgress(channel.Writer);
        var task = Task.Run(async () =>
        {
            try { return await work(progress); }
            catch (Exception ex) when (ex is not ReportFailure and not OperationCanceledException)
            {
                progress.Report($"[ERROR] {ex.Message}");
                throw;
            }
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
        catch (ReportFailure ex) { await Write(new { type = "error", message = ex.Message, panel = ex.Panel }); }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            await Write(new { type = "error", message = $"Error: {ex.Message}", panel = false });
        }
    }

    private static IResult Error(string message) => Results.Json(new { error = message }, statusCode: StatusCodes.Status400BadRequest);

    public static IEndpointRouteBuilder MapNgRate(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/ngrate").RequireAuthorization();
        api.AddEndpointFilter(async (ctx, next) =>
        {
            var http = ctx.HttpContext;
            if (!http.RequestServices.GetRequiredService<MenuPermissionService>().IsAllowed(http.User, AppMenus.NgRate))
                return Results.StatusCode(StatusCodes.Status403Forbidden);
            // Changes must come from the React app (custom header → browsers preflight cross-site requests).
            if (!HttpMethods.IsGet(http.Request.Method) &&
                http.Request.Headers[AdminUserEndpoints.RequestedWithHeader] != AdminUserEndpoints.RequestedWithValue)
                return Error("Required request header is missing.");
            return await next(ctx);
        });

        // Credentials state + the Model Groups picker (both child-entry flavours).
        api.MapGet("/setup", (WebRepository repo, NgRateSettingsService settings) => Results.Ok(Setup(repo, settings)));

        // By Model / By Section report run (NgRatePage / NgRateAllPage GetReportAsync).
        api.MapPost("/report", async (
            HttpContext ctx, ReportRequest body, NgRateService ngRate, NgRateSettingsService settings,
            NgRateReportService reportSvc, WebRepository repo, AppActivityLogger activity) =>
        {
            bool bySection = string.Equals(body.Mode, "by-section", StringComparison.Ordinal);
            if (!bySection && !string.Equals(body.Mode, "by-model", StringComparison.Ordinal))
                return Error("Unknown report mode.");
            if (!TryParseDay(body.Start, out DateTime start) || !TryParseDay(body.End, out DateTime end))
                return Error("Enter the period as yyyy-MM-dd.");
            if (end < start) return Error("End date must be on or after start date.");
            if (!settings.IsCredentialsConfigured) return Error("BMES credentials are not configured.");

            string area = bySection ? AreaBySection : AreaByModel;
            string sep = bySection ? "---" : "───";
            Func<ModelGroupRecord, IEnumerable<ModelGroupChildEntry>> childEntries = bySection
                ? ModelGroupPickerHelpers.EnumerateMidEntries
                : ModelGroupPickerHelpers.EnumerateChildEntries;
            activity.Log(area, $"Get Data {start:yyyy-MM-dd} ~ {end:yyyy-MM-dd}");

            await Stream(ctx, async progress =>
            {
                progress.Report($"Fetching data: {start:yyyy-MM-dd} ~ {end:yyyy-MM-dd}");
                string? dbPath = await ngRate.GetOrFetchAsync(start, end, progress);
                if (string.IsNullOrEmpty(dbPath))
                    throw new ReportFailure("Data fetch or save failed - check the log.", panel: true);

                // Same as the page's RefreshGroups: picks that no longer exist are dropped.
                var groups = repo.GetModelGroups();
                var groupIds = (body.GroupIds ?? []).Where(id => groups.Any(g => g.Id == id)).ToHashSet();
                var live = groups.SelectMany(childEntries).Select(c => c.Key).ToHashSet(StringComparer.Ordinal);
                var subEntries = (body.SubEntries ?? []).Where(live.Contains).Distinct(StringComparer.Ordinal).ToList();

                activity.Log(area, $"Generate Report selectedGroups={groupIds.Count:N0}, selectedEntries={subEntries.Count:N0}");
                progress.Report($"Using NG data: {dbPath}");

                var reports = new List<(string Label, NgRateReportService.NgRateReport Report)>();
                if (groupIds.Count == 0 && subEntries.Count == 0)
                {
                    // No selection → single "All" report across every group.
                    var flat = NgRateModeSupport.BuildFlatMapping(groups);
                    if (flat.Mapping.Count == 0) throw new ReportFailure("No LineShift found in the selected groups.");
                    reports.Add(("All", await reportSvc.GenerateReportAsync(dbPath, flat.Mapping, flat.Groups, progress, start, end)));
                }
                else
                {
                    foreach (var g in groups.Where(g => groupIds.Contains(g.Id)).OrderBy(g => g.Name, StringComparer.Ordinal))
                    {
                        progress.Report($"{sep} {g.Name}");
                        if (bySection)
                        {
                            var flat = NgRateModeSupport.BuildFlatMapping([g]);
                            if (flat.Mapping.Count == 0) { progress.Report("  (skipped: no LineShift)"); continue; }
                            reports.Add((g.Name, await reportSvc.GenerateReportAsync(dbPath, flat.Mapping, flat.Groups, progress, start, end)));
                        }
                        else
                        {
                            var tree = NgRateModeSupport.BuildTreeMultiMapping(g);
                            if (tree.Mapping.Count == 0) { progress.Report("  (skipped: no LineShift)"); continue; }
                            reports.Add((g.Name, await reportSvc.GenerateReportAsync(dbPath, tree.Mapping, tree.Groups, progress, start, end)));
                        }
                    }

                    foreach (string key in subEntries.OrderBy(k => k, StringComparer.Ordinal))
                    {
                        if (bySection)
                        {
                            var spec = NgRateModeSupport.BuildMidChildFlatReport(groups, key);
                            if (spec is null) { progress.Report($"{sep} {key}  (skipped: not found)"); continue; }
                            progress.Report($"{sep} {spec.Label}");
                            reports.Add((spec.Label, await reportSvc.GenerateReportAsync(dbPath, spec.Mapping, spec.Groups, progress, start, end)));
                        }
                        else
                        {
                            var spec = NgRateModeSupport.BuildChildTreeReport(groups, key);
                            if (spec is null) { progress.Report($"{sep} {key}  (skipped: not found)"); continue; }
                            progress.Report($"{sep} {spec.Label}");
                            reports.Add((spec.Label, await reportSvc.GenerateReportAsync(dbPath, spec.Mapping, spec.Groups, progress, start, end)));
                        }
                    }

                    if (!bySection && reports.Count == 0)
                        throw new ReportFailure("No LineShift found in the selected groups.");
                }

                // Group → Model → Sub → Shift card over the union of the selected scopes.
                // (By Model built it over the whole fetched range; By Section over the picked period.)
                var hierGroups = ResolveHierGroups(groups, groupIds, subEntries);
                var hier = bySection
                    ? await HierReportBuilder.BuildAsync(reportSvc, dbPath, hierGroups, progress, start, end)
                    : await HierReportBuilder.BuildAsync(reportSvc, dbPath, hierGroups, progress);

                string? exportId = reports.Count == 0
                    ? null
                    : Remember(new ExportEntry(UserOf(ctx), DateTime.Now, area, "NgRate", reports, null));
                return new ReportResult(
                    dbPath,
                    reports.Select(r => new LabeledReportDto(r.Label, ToDto(r.Report))).ToList(),
                    ToHier(hier),
                    exportId);
            });
            return Results.Empty;
        });

        // By Group report run (NgRateByGroupPage GetReportAsync).
        api.MapPost("/by-group", async (
            HttpContext ctx, ByGroupRequest body, NgRateService ngRate, NgRateSettingsService settings,
            NgRateReportService reportSvc, WebRepository repo, AppActivityLogger activity) =>
        {
            var names = (body.GroupNames ?? []).Where(n => !string.IsNullOrEmpty(n)).Distinct(StringComparer.Ordinal).ToList();
            if (names.Count == 0) return Error("Select at least one group.");
            if (!TryParseDay(body.Start, out DateTime start) || !TryParseDay(body.End, out DateTime end))
                return Error("Enter the period as yyyy-MM-dd.");
            if (end < start) return Error("End date must be on or after start date.");
            if (!settings.IsCredentialsConfigured) return Error("BMES credentials are not configured.");

            activity.Log(AreaByGroup, $"Get Data {start:yyyy-MM-dd} ~ {end:yyyy-MM-dd}");
            await Stream(ctx, async progress =>
            {
                progress.Report($"Fetching data: {start:yyyy-MM-dd} ~ {end:yyyy-MM-dd}");
                string? dbPath = await ngRate.GetOrFetchAsync(start, end, progress);
                if (string.IsNullOrEmpty(dbPath))
                    throw new ReportFailure("Data fetch or save failed - check the log.", panel: true);

                var groups = repo.GetModelGroups();
                var picked = names.Where(n => groups.Any(g => g.Name == n)).ToList();
                activity.Log(AreaByGroup, $"Generate Report selectedGroups={picked.Count:N0}");
                progress.Report($"Using NG data: {dbPath}");

                var selected = NgRateModeSupport.SelectGroupsByName(groups, picked.ToHashSet(StringComparer.Ordinal));
                var bundle = await NgRateModeSupport.GenerateGroupReportBundleAsync(reportSvc, dbPath, selected, progress, start, end)
                    ?? throw new ReportFailure("No LineShift found in the selected groups.");

                var reports = new List<(string Label, NgRateReportService.NgRateReport Report)>
                {
                    ("By Group - Group", bundle.GroupReport),
                    ("By Group - Model", bundle.MidReport),
                    ("By Group - Shift", bundle.LineShiftReport),
                };
                if (bundle.Sub1Report is not null) reports.Add(("By Group - Sub Group", bundle.Sub1Report));
                if (bundle.Sub2Report is not null) reports.Add(("By Group - Shift Group", bundle.Sub2Report));
                var details = new List<(string Label, NgRateReportService.LineShiftNgReport Report)>
                {
                    ("By Group Detail - Model", bundle.MidNgReport),
                    ("By Group Detail - Shift", bundle.LineShiftNgReport),
                };
                string exportId = Remember(new ExportEntry(UserOf(ctx), DateTime.Now, AreaByGroup, "NgRate_ByGroup", reports, details));
                return ToByGroup(bundle, groups, picked, dbPath, exportId);
            });
            return Results.Empty;
        });

        // Workbook of a finished run (NgRateExcelExporter, same sheets as the Blazor Export Excel).
        api.MapGet("/export/{id}", (HttpContext ctx, string id, AppActivityLogger activity) =>
        {
            if (!Exports.TryGetValue(id, out var entry) || entry.User != UserOf(ctx))
                return Results.NotFound(new { error = "This report is no longer available. Run Get Report again." });
            activity.Log(entry.Area, $"Export Excel reports={entry.Reports.Count:N0}");
            byte[] bytes = NgRateExcelExporter.Export(entry.Reports, entry.Details);
            return Results.File(bytes, BrowserDownload.XlsxContentType, $"{entry.FilePrefix}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        });

        return app;
    }
}
