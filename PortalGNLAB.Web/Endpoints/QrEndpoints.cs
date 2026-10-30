using System.Globalization;
using PortalGNLAB.Web.Services;

namespace PortalGNLAB.Web.Endpoints;

/// <summary>
/// QR pages for the React app (/app/qr/summary, /app/qr/*-data, /app/qr/gauss-scatter): port of
/// Components/Pages/QrSummaryPage.razor, Components/Shared/QrTestDataBrowser.razor and
/// Components/Pages/QrGaussScatterPage.razor. The existing services do the TCDB reads (SELECT only);
/// this file only shapes their results as JSON and checks the same menu permissions as the nav.
/// </summary>
public static class QrEndpoints
{
    public const string XlsxContentType = QrTestDataExcelExporter.ContentType;

    /// <summary>Response header set on a workbook that stopped at the export ceiling.</summary>
    public const string TruncatedHeader = "X-Qr-Truncated";

    public sealed record QrConnection(string Server, int Port, string Database, string UserId);

    public sealed record QrModelCodeDto(string Code, string Name);

    public sealed record QrPrefsView(
        bool Configured,
        QrConnection Connection,
        int DefaultMaxRows,
        int MaxRows,
        int MaxQr,
        IReadOnlyList<QrModelCodeDto> ModelCodes);

    public sealed record QrPrefsChange(int? MaxRows, int? MaxQr);

    public sealed record QrItemDto(string Label, string Value, string Verdict);

    public sealed record QrCellDto(string Station, string Verdict, string Result, string? Time, IReadOnlyList<QrItemDto> Items);

    public sealed record QrSummaryRowDto(
        string ProductId,
        int Round,
        string? LatestTime,
        bool FinalPass,
        string FinalNote,
        bool HasAnyFail,
        bool Complete,
        QrCellDto? Airleak,
        QrCellDto? Bako,
        QrCellDto? Gauss,
        IReadOnlyList<QrSummaryRowDto> Older);

    public sealed record QrSummaryView(
        string From,
        string To,
        int MaxRows,
        long TotalProducts,
        bool HitMaxRows,
        string? MaxDate,
        string FetchedAt,
        string Warning,
        IReadOnlyList<QrSummaryRowDto> Rows);

    public sealed record QrColumnDto(string Name, string DataType, bool Numeric);

    public sealed record QrDataRowDto(IReadOnlyList<string> Values, string? Date);

    public sealed record QrDateDto(string Date, long RowCount);

    public sealed record QrSummaryCardDto(long InputCount, long NgCount, double? NgRatePpm);

    public sealed record QrTableDto(
        string Key,
        string Station,
        string MenuLabel,
        string DisplayTable,
        IReadOnlyList<string> PassValues,
        IReadOnlyList<string> FailValues);

    public sealed record QrDataView(
        QrTableDto Table,
        string FetchedAt,
        int MaxRows,
        long? TotalRows,
        long? ValidTotalRows,
        long? ExcludedRows,
        string SortColumnName,
        IReadOnlyList<string> SelectedDates,
        string Warning,
        bool HitMaxRows,
        QrSummaryCardDto? Summary,
        IReadOnlyList<QrDateDto> Dates,
        IReadOnlyList<QrColumnDto> Columns,
        IReadOnlyList<int> Layout,
        IReadOnlyList<QrDataRowDto> Rows);

    public sealed record GaussPointDto(string ProductId, string Time, string Date, double Value, bool Ng, double Jitter);

    public sealed record GaussScatterView(
        string From,
        string To,
        int MaxRows,
        string? MaxDate,
        string FetchedAt,
        string Warning,
        IReadOnlyList<string> ProductIds,
        IReadOnlyList<GaussPointDto> Value,
        IReadOnlyList<GaussPointDto> Value2,
        int RawValueCount,
        int RawValue2Count,
        int NgValueCount,
        int NgValue2Count);

    private static string? Day(DateOnly? date) => date?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static string Day(DateOnly date) => date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static string Stamp(DateTimeOffset time) => time.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

    private static string Verdict(QrVerdict verdict) => verdict switch
    {
        QrVerdict.Pass => "pass",
        QrVerdict.Fail => "fail",
        _ => "unknown",
    };

    public static bool TryParseDay(string? text, out DateOnly date) =>
        DateOnly.TryParseExact(text?.Trim(), "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out date);

    public static List<DateOnly> ParseDays(string? csv)
    {
        var dates = new List<DateOnly>();
        foreach (string part in (csv ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (TryParseDay(part, out DateOnly date) && !dates.Contains(date))
                dates.Add(date);
        }
        return dates;
    }

    private static QrCellDto? ToDto(QrSummaryCell? cell) => cell is null
        ? null
        : new QrCellDto(
            cell.Station,
            Verdict(cell.Verdict),
            cell.Result,
            cell.TestTime is DateTimeOffset time ? Stamp(time) : null,
            cell.Items.Select(item => new QrItemDto(item.Label, item.Value, Verdict(item.Verdict))).ToList());

    private static QrSummaryRowDto ToDto(QrSummaryRow row, bool withOlder) => new(
        row.ProductId,
        row.Round,
        row.LatestTime is DateTimeOffset latest ? Stamp(latest) : null,
        row.IsFinalPass,
        row.FinalNote,
        row.HasAnyFail,
        row.IsCompleteSet,
        ToDto(row.Airleak),
        ToDto(row.Bako),
        ToDto(row.Gauss),
        withOlder ? row.Older.Select(older => ToDto(older, false)).ToList() : []);

    public static QrSummaryView ToView(QrSummaryResult result) => new(
        Day(result.From),
        Day(result.To),
        result.MaxRows,
        result.TotalProducts,
        result.HitMaxRows,
        Day(result.Bounds.MaxDate),
        result.FetchedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        result.WarningMessage,
        result.Rows.Select(row => ToDto(row, true)).ToList());

    public static QrDataView ToView(QrTestTableDef table, QrTestDataResult result) => new(
        new QrTableDto(table.Key, table.Station, table.MenuLabel, table.DisplayTable, table.PassValues, table.FailValues),
        result.FetchedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        result.MaxRows,
        result.TotalRows,
        result.ValidTotalRows,
        result.ExcludedRows,
        result.SortColumnName,
        result.SelectedDates.Select(d => Day(d)).ToList(),
        result.WarningMessage,
        result.HitMaxRows,
        result.Summary is { } summary ? new QrSummaryCardDto(summary.InputCount, summary.NgCount, summary.NgRatePpm) : null,
        result.Dates.Select(d => new QrDateDto(Day(d.Date), d.RowCount)).ToList(),
        result.Columns.Select(c => new QrColumnDto(c.Name, c.DataType, QrTestColumnLayout.IsNumeric(c.DataType))).ToList(),
        QrTestColumnLayout.Build(table, result.Columns),
        result.Rows.Select(row => new QrDataRowDto(row.Values, Day(row.TestDate))).ToList());

    /// <summary>
    /// A repeatable −1..1 spread for one reading (FNV-1a over the QR and its timestamp), so a point stays
    /// where it is across redraws. Same arithmetic as the Blazor page; <c>string.GetHashCode</c> is per-process.
    /// </summary>
    public static double Jitter(string productId, DateTimeOffset time, int slot)
    {
        unchecked
        {
            uint hash = 2166136261;
            foreach (char c in productId)
                hash = (hash ^ c) * 16777619;

            hash = (hash ^ (uint)time.Ticks) * 16777619;
            hash = (hash ^ (uint)(time.Ticks >> 32)) * 16777619;
            hash = (hash ^ (uint)slot) * 16777619;
            return hash / (double)uint.MaxValue * 2 - 1;
        }
    }

    /// <summary>
    /// GSTD readings arrive as text and a tester occasionally appends a unit: take the leading
    /// numeric token and ignore the rest.
    /// </summary>
    public static bool TryParseNumber(string? text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text))
            return false;

        ReadOnlySpan<char> span = text.AsSpan().Trim();
        int end = 0;
        while (end < span.Length &&
               (char.IsAsciiDigit(span[end]) || span[end] is '+' or '-' or '.' or 'e' or 'E'))
        {
            end++;
        }

        return double.TryParse(span[..end], NumberStyles.Float, CultureInfo.InvariantCulture, out value)
            && !double.IsNaN(value)
            && !double.IsInfinity(value);
    }

    public static GaussScatterView ToScatter(QrSummaryResult result)
    {
        var value = new List<(DateTimeOffset Time, GaussPointDto Point)>();
        var value2 = new List<(DateTimeOffset Time, GaussPointDto Point)>();
        int ngValue = 0, ngValue2 = 0;

        // Every round counts, not just the newest: dispersion is exactly what the retests show.
        foreach (QrSummaryRow product in result.Rows)
        {
            Collect(product.ProductId, product);
            foreach (QrSummaryRow older in product.Older)
                Collect(product.ProductId, older);
        }

        void Collect(string productId, QrSummaryRow row)
        {
            if (row.Gauss is not QrSummaryCell cell || cell.TestTime is not DateTimeOffset time)
                return;

            bool ng = cell.Verdict == QrVerdict.Fail;
            string date = Day(DateOnly.FromDateTime(time.LocalDateTime));
            Add(value, cell.ItemAt(0), 0, ref ngValue);
            Add(value2, cell.ItemAt(1), 1, ref ngValue2);

            void Add(List<(DateTimeOffset, GaussPointDto)> target, QrSummaryItem? item, int slot, ref int ngCount)
            {
                if (item is null || !TryParseNumber(item.Value, out double parsed))
                    return;
                target.Add((time, new GaussPointDto(productId, Stamp(time), date, parsed, ng, Jitter(productId, time, slot))));
                if (ng)
                    ngCount++;
            }
        }

        static List<GaussPointDto> Ordered(List<(DateTimeOffset Time, GaussPointDto Point)> points) =>
            points.OrderBy(p => p.Time).Select(p => p.Point).ToList();

        return new GaussScatterView(
            Day(result.From),
            Day(result.To),
            result.MaxRows,
            Day(result.Bounds.MaxDate),
            result.FetchedAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
            result.WarningMessage,
            result.Rows.Select(r => r.ProductId).ToList(),
            Ordered(value),
            Ordered(value2),
            value.Count,
            value2.Count,
            ngValue,
            ngValue2);
    }

    /// <summary>Menu a data table belongs to; the same ids the Blazor nav gates the pages with.</summary>
    public static string? MenuForTable(string key) => key.ToUpperInvariant() switch
    {
        QrTestTables.BakoKey => AppMenus.QrBakoData,
        QrTestTables.AirleakKey => AppMenus.QrAirleakData,
        QrTestTables.GaussKey => AppMenus.QrGaussData,
        _ => null,
    };

    private static readonly string[] AllQrMenus =
        [AppMenus.QrSummary, AppMenus.QrBakoData, AppMenus.QrAirleakData, AppMenus.QrGaussData];

    private static IResult Fail(Exception ex) =>
        Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);

    public static IEndpointRouteBuilder MapQr(this IEndpointRouteBuilder app)
    {
        var api = app.MapGroup("/api/qr").RequireAuthorization();

        // Changes must come from the React app (custom header → browsers preflight cross-site requests).
        api.AddEndpointFilter(async (ctx, next) =>
        {
            var request = ctx.HttpContext.Request;
            if (!HttpMethods.IsGet(request.Method) &&
                request.Headers[AdminUserEndpoints.RequestedWithHeader] != AdminUserEndpoints.RequestedWithValue)
                return Results.StatusCode(StatusCodes.Status400BadRequest);
            return await next(ctx);
        });

        static bool Allowed(HttpContext ctx, params string[] menus)
        {
            var perms = ctx.RequestServices.GetRequiredService<MenuPermissionService>();
            return menus.Any(menu => perms.IsAllowed(ctx.User, menu));
        }

        // What the pages remember between visits, plus the model-code table and the connection label.
        api.MapGet("/prefs", (HttpContext ctx, QrTestDataService data, QrViewPreferenceService prefs, QrModelCodeService models) =>
        {
            if (!Allowed(ctx, AllQrMenus)) return Results.Forbid();
            QrTestDataConnectionSnapshot snapshot = data.Snapshot;
            return Results.Ok(new QrPrefsView(
                data.IsConfigured,
                new QrConnection(snapshot.Server, snapshot.Port, snapshot.Database, snapshot.UserId),
                data.DefaultMaxRows,
                prefs.MaxRows,
                prefs.MaxQr,
                models.Codes.Select(c => new QrModelCodeDto(c.Code, c.Name)).ToList()));
        });

        api.MapPut("/prefs", (HttpContext ctx, QrPrefsChange body, QrViewPreferenceService prefs) =>
        {
            if (!Allowed(ctx, AllQrMenus)) return Results.Forbid();
            if (body.MaxRows is int rows) prefs.SetMaxRows(rows);
            if (body.MaxQr is int qr) prefs.SetMaxQr(qr);
            return Results.NoContent();
        });

        api.MapGet("/summary", async (HttpContext ctx, string? from, string? to, int? max, QrSummaryService summary, CancellationToken ct) =>
        {
            if (!Allowed(ctx, AppMenus.QrSummary)) return Results.Forbid();
            try
            {
                QrSummaryResult result = await summary.FetchAsync(new QrSummaryQuery
                {
                    From = TryParseDay(from, out DateOnly f) ? f : default,
                    To = TryParseDay(to, out DateOnly t) ? t : default,
                    MaxRows = max ?? 2000,
                }, ct);
                return Results.Ok(ToView(result));
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { return Fail(ex); }
        });

        // The workbook comes from its own unlimited read of the period, not from what the grid shows.
        api.MapGet("/summary/export", async (HttpContext ctx, string? from, string? to, QrSummaryService summary, CancellationToken ct) =>
        {
            if (!Allowed(ctx, AppMenus.QrSummary)) return Results.Forbid();
            try
            {
                QrSummaryResult full = await summary.FetchAsync(new QrSummaryQuery
                {
                    From = TryParseDay(from, out DateOnly f) ? f : default,
                    To = TryParseDay(to, out DateOnly t) ? t : default,
                    AllRows = true,
                }, ct);
                ctx.Response.Headers[TruncatedHeader] = full.Rows.Count >= QrSummaryService.ExportMaxRows ? "1" : "0";
                return Results.File(QrSummaryExcelExporter.Build(full.Rows), QrSummaryExcelExporter.ContentType, "QR_SUMMARY.xlsx");
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { return Fail(ex); }
        });

        // Dates arrive as a comma list; maxRows 0 means "fit the selected dates".
        api.MapGet("/data/{key}", async (HttpContext ctx, string key, string? dates, int? maxRows, QrTestDataService data, CancellationToken ct) =>
        {
            string? menu = MenuForTable(key);
            if (menu is null) return Results.NotFound();
            if (!Allowed(ctx, menu)) return Results.Forbid();
            QrTestTableDef table = QrTestTables.Resolve(key);
            try
            {
                QrTestDataResult result = await data.FetchAsync(new QrTestDataQuery
                {
                    Table = table,
                    MaxRows = maxRows is > 0 ? maxRows.Value : data.DefaultMaxRows,
                    AutoMaxRows = maxRows is null or <= 0,
                    TestDates = ParseDays(dates),
                }, ct);
                return Results.Ok(ToView(table, result));
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { return Fail(ex); }
        });

        api.MapGet("/data/{key}/export", async (HttpContext ctx, string key, string? dates, QrTestDataService data, CancellationToken ct) =>
        {
            string? menu = MenuForTable(key);
            if (menu is null) return Results.NotFound();
            if (!Allowed(ctx, menu)) return Results.Forbid();
            QrTestTableDef table = QrTestTables.Resolve(key);
            try
            {
                QrTestDataResult full = await data.FetchAsync(new QrTestDataQuery
                {
                    Table = table,
                    AllRows = true,
                    TestDates = ParseDays(dates),
                }, ct);
                ctx.Response.Headers[TruncatedHeader] = full.Rows.Count >= QrTestDataService.ExportMaxRows ? "1" : "0";
                return Results.File(QrTestDataExcelExporter.Build(table, full), XlsxContentType, $"QR_{table.Station}_DATA.xlsx");
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { return Fail(ex); }
        });

        api.MapGet("/gauss-scatter", async (HttpContext ctx, string? from, string? to, int? max, QrSummaryService summary, CancellationToken ct) =>
        {
            if (!Allowed(ctx, AppMenus.QrSummary, AppMenus.QrGaussData)) return Results.Forbid();
            try
            {
                QrSummaryResult result = await summary.FetchAsync(new QrSummaryQuery
                {
                    From = TryParseDay(from, out DateOnly f) ? f : default,
                    To = TryParseDay(to, out DateOnly t) ? t : default,
                    MaxRows = max ?? 2000,
                }, ct);
                return Results.Ok(ToScatter(result));
            }
            catch (Exception ex) when (ex is not OperationCanceledException) { return Fail(ex); }
        });

        return app;
    }
}
