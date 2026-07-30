using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// MES050032 IPG defect response client/parser.
/// Authentication always starts from configured credentials and a new cookie session;
/// copied browser cookies and SSO/session tokens are not inputs to this service.
/// </summary>
public sealed class IpgDefectService(
    NgRateSettingsService settings,
    AppActivityLogger activity)
{
    private const string BaseUrl = "https://bmes.bujeon.com";
    private readonly NgRateSettingsService _settings = settings;
    private readonly AppActivityLogger _activity = activity;

    public async Task<IpgDefectParseResult> FetchAsync(
        DateTime queryDate,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        using HttpClient client = CreateClient();
        await AuthenticateAsync(client, progress, cancellationToken);
        return await FetchSearchListAsync(
            client,
            queryDate,
            "W",
            progress,
            cancellationToken);
    }

    public async Task<IpgDefectKpiSnapshot> FetchKpiRangeAsync(
        DateTime start,
        DateTime end,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        start = start.Date;
        end = end.Date;
        if (end > DateTime.Today) end = DateTime.Today;
        if (start > end)
            throw new ArgumentException("Start date must be before or equal to end date.");

        using HttpClient client = CreateClient();
        await AuthenticateAsync(client, progress, cancellationToken);

        var periodsByKey =
            new Dictionary<string, IpgDefectKpiPeriod>(StringComparer.OrdinalIgnoreCase);
        double? annualAveragePpm = null;
        List<DateTime> weeklyQueryDates = BuildQueryDates(start, end, intervalDays: 42);

        for (int index = 0; index < weeklyQueryDates.Count; index++)
        {
            DateTime queryDate = weeklyQueryDates[index];
            try
            {
                Report(
                    progress,
                    $"MES050032 weekly [{index + 1}/{weeklyQueryDates.Count}]: {queryDate:yyyy-MM-dd}");
                IpgDefectParseResult parsed = await FetchSearchListAsync(
                    client,
                    queryDate,
                    "W",
                    progress,
                    cancellationToken);
                IpgDefectKpiSnapshot snapshot = BuildKpiSnapshot(parsed, "Week");
                annualAveragePpm = snapshot.AnnualAveragePpm;
                foreach (IpgDefectKpiPeriod period in snapshot.Periods)
                    periodsByKey[IpgPeriodIdentity(period)] = period;
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                Report(
                    progress,
                    $"[WARN] MES050032 weekly {queryDate:yyyy-MM-dd}: {ex.Message}");
            }
        }

        try
        {
            Report(progress, $"MES050032 monthly: {end:yyyy-MM-dd}");
            IpgDefectParseResult parsed = await FetchSearchListAsync(
                client,
                end,
                "M",
                progress,
                cancellationToken);
            IpgDefectKpiSnapshot monthly = BuildKpiSnapshot(parsed, "Month");
            foreach (IpgDefectKpiPeriod period in monthly.Periods)
                periodsByKey[IpgPeriodIdentity(period)] = period;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // A monthly response is optional until BMES actually returns trustworthy
            // header metadata. No inferred monthly value is generated on failure.
            Report(progress, "[WARN] MES050032 monthly: " + ex.Message);
        }

        return new IpgDefectKpiSnapshot(
            periodsByKey.Values.ToList(),
            annualAveragePpm);
    }

    public static string BuildSearchListPayload(DateTime queryDate)
        => BuildSearchListPayload(queryDate, "W");

    public static string BuildSearchListPayload(DateTime queryDate, string periodCode)
    {
        string normalizedPeriodCode = NormalizePeriodCode(periodCode);
        var payload = new
        {
            Condition = new
            {
                SDATE = queryDate.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                ZGUBN = normalizedPeriodCode,
                V_FACCO = (string?)null,
                V_MATNR = (string?)null,
                V_WERKS = (string?)null,
                V_VERID = (string?)null,
                ZLRTP = "L",
                ZTYPE = "A",
                V_MAINS = "A",
                V_COLNM = "COL0001",
                V_CPTYP = "J02",
            },
            L_FACCO = new[] { "GN" },
            L_DIVIS = new[] { "" },
            L_CATEG = new[] { "" },
            L_PRODU = new[] { "" },
            L_ITEMS = new[] { "" },
            L_MODEL = new[] { "" },
            L_SERIE = new[] { "" },
            L_MODEL2 = new[] { "" },
            L_VERID = new[] { "" },
            L_MATNR = new[] { "" },
            REMEM = "N",
        };
        return JsonSerializer.Serialize(payload);
    }

    public static IpgDefectParseResult ParseSearchListJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        string resultCode = ReadText(root, "Result");
        if (resultCode.Length > 0 &&
            !string.Equals(resultCode, "S", StringComparison.OrdinalIgnoreCase))
        {
            string message = ReadText(root, "Msg");
            throw new JsonException(
                message.Length > 0
                    ? $"MES050032 returned {resultCode}: {message}"
                    : $"MES050032 returned {resultCode}.");
        }

        var columns = new List<IpgDefectColumn>();
        if (TryGetProperty(root, "header", out JsonElement header) &&
            header.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement item in header.EnumerateArray())
            {
                string columnName = ReadText(item, "ZCOLN");
                int index = ParseColumnIndex(columnName);
                if (index is < 1 or > 14)
                    continue;

                columns.Add(new IpgDefectColumn(
                    index,
                    columnName,
                    ReadText(item, "ZCOLT")));
            }
        }

        if (!TryGetProperty(root, "ChartRec", out JsonElement chartRecords) ||
            chartRecords.ValueKind != JsonValueKind.Array)
        {
            throw new JsonException("MES050032 response does not contain ChartRec.");
        }

        var rows = new List<IpgDefectRow>(chartRecords.GetArrayLength());
        int observedColumnCount = 0;
        foreach (JsonElement item in chartRecords.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
                continue;

            var row = new IpgDefectRow
            {
                RawJson = item.GetRawText(),
                Cptyp = ReadText(item, "CPTYP"),
                CptypTx = ReadText(item, "CPTYP_TX"),
                FaccoTx = ReadText(item, "FACCO_TX"),
                ModnoTx = ReadText(item, "MODNO_TX"),
                ZlrtpTx = ReadText(item, "ZLRTP_TX"),
                MatnrTx = ReadText(item, "MATNR_TX"),
                Facco = ReadText(item, "FACCO"),
                Modno = ReadText(item, "MODNO"),
                Zsort = ReadText(item, "ZSORT"),
                Zlrtp = ReadText(item, "ZLRTP"),
                Avegr = ReadNullableDouble(item, "AVEGR"),
            };

            for (int index = 1; index <= 14; index++)
            {
                string propertyName = $"COL{index:D4}";
                if (TryGetProperty(item, propertyName, out JsonElement value))
                {
                    observedColumnCount = Math.Max(observedColumnCount, index);
                    row.SetCol(index, ReadNullableDouble(value));
                }
            }

            rows.Add(row);
        }

        columns = columns
            .GroupBy(column => column.ColIndex)
            .Select(group => group.First())
            .OrderBy(column => column.ColIndex)
            .ToList();

        int columnCount = Math.Max(
            observedColumnCount,
            columns.Count > 0 ? columns.Max(column => column.ColIndex) : 0);
        return new IpgDefectParseResult(rows, columns, columnCount);
    }

    public static IpgDefectKpiSnapshot BuildKpiSnapshot(
        IpgDefectParseResult parsed,
        string periodKind = "Week")
    {
        ArgumentNullException.ThrowIfNull(parsed);
        string normalizedPeriodKind = NormalizePeriodKind(periodKind);

        IReadOnlyList<IpgDefectKpiPeriod> periods = parsed.Columns
            .OrderBy(column => column.ColIndex)
            .Select(column =>
            {
                double[] values = parsed.Rows
                    .Select(row => row.GetCol(column.ColIndex))
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value)
                    .ToArray();
                return new IpgDefectKpiPeriod(
                    column.ColIndex,
                    column.ColumnName,
                    column.Header,
                    normalizedPeriodKind,
                    values.Length > 0 ? values.Average() : null);
            })
            .ToList();

        double[] annualValues = parsed.Rows
            .Where(row => row.Avegr.HasValue)
            .Select(row => row.Avegr!.Value)
            .ToArray();
        return new IpgDefectKpiSnapshot(
            periods,
            annualValues.Length > 0 ? annualValues.Average() : null);
    }

    private static HttpClient CreateClient()
    {
        var handler = new HttpClientHandler
        {
            UseCookies = true,
            CookieContainer = new System.Net.CookieContainer(),
        };
        var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(300) };
        client.DefaultRequestHeaders.Add("X-Requested-With", "XMLHttpRequest");
        return client;
    }

    private async Task AuthenticateAsync(
        HttpClient client,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string loginId = _settings.LoginId;
        string password = _settings.Password;
        if (string.IsNullOrWhiteSpace(loginId) || string.IsNullOrWhiteSpace(password))
            throw new InvalidOperationException("BMES credentials are not configured.");

        Report(progress, "MES050032: reading login token");
        string token = await GetTokenAsync(client, cancellationToken);
        if (string.IsNullOrEmpty(token))
            throw new InvalidOperationException("Failed to read BMES login token.");

        Report(progress, "MES050032: logging in with configured credentials");
        if (!await LoginAsync(client, token, loginId, password, cancellationToken))
            throw new InvalidOperationException("BMES login failed.");
    }

    private async Task<IpgDefectParseResult> FetchSearchListAsync(
        HttpClient client,
        DateTime queryDate,
        string periodCode,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string normalizedPeriodCode = NormalizePeriodCode(periodCode);
        string payload = BuildSearchListPayload(queryDate, normalizedPeriodCode);
        using var content = new StringContent(payload, Encoding.UTF8, "application/json");
        using HttpResponseMessage response = await client.PostAsync(
            BaseUrl + "/MES050032/SearchList",
            content,
            cancellationToken);
        string json = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();

        IpgDefectParseResult parsed = ParseSearchListJson(json);
        Report(
            progress,
            $"MES050032 {normalizedPeriodCode}: parsed {parsed.Rows.Count:N0} rows · {parsed.Columns.Count:N0} headers");
        return parsed;
    }

    private static List<DateTime> BuildQueryDates(
        DateTime start,
        DateTime end,
        int intervalDays)
    {
        var dates = new HashSet<DateTime> { start.Date, end.Date };
        for (DateTime cursor = end.Date;
             cursor > start.Date && cursor >= DateTime.MinValue.AddDays(intervalDays);
             cursor = cursor.AddDays(-intervalDays))
        {
            dates.Add(cursor);
        }
        return dates.OrderBy(date => date).ToList();
    }

    private static string NormalizePeriodCode(string periodCode) =>
        (periodCode ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "W" or "WEEK" => "W",
            "M" or "MONTH" => "M",
            _ => throw new ArgumentException(
                "MES050032 period must be W (week) or M (month).",
                nameof(periodCode)),
        };

    private static string NormalizePeriodKind(string periodKind) =>
        NormalizePeriodCode(periodKind) == "M" ? "Month" : "Week";

    private static string IpgPeriodIdentity(IpgDefectKpiPeriod period)
    {
        string periodCode = string.IsNullOrWhiteSpace(period.Header)
            ? period.ColumnName
            : period.Header;
        return period.Kind + "|" + periodCode;
    }

    private void Report(IProgress<string>? progress, string message)
    {
        _activity.Log("IpgDefect", message);
        progress?.Report(message);
    }

    private static async Task<string> GetTokenAsync(
        HttpClient client,
        CancellationToken cancellationToken)
    {
        try
        {
            string html = await client.GetStringAsync(BaseUrl, cancellationToken);
            Match match = Regex.Match(
                html,
                @"<input[^>]+name=""__RequestVerificationToken""[^>]+value=""([^""]+)""",
                RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                match = Regex.Match(
                    html,
                    @"<input[^>]+value=""([^""]+)""[^>]+name=""__RequestVerificationToken""",
                    RegexOptions.IgnoreCase);
            }
            return match.Success ? match.Groups[1].Value : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static async Task<bool> LoginAsync(
        HttpClient client,
        string token,
        string loginId,
        string password,
        CancellationToken cancellationToken)
    {
        using var content = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["UserInfo[USRID]"] = loginId,
            ["UserInfo[PWNO]"] = password,
            ["UserInfo[LANG]"] = "EN",
            ["UserInfo[FACCO]"] = "GN",
            ["UserInfo[STYPE]"] = "P",
            ["UserInfo[VTYPE]"] = "P",
            ["__RequestVerificationToken"] = token,
        });
        try
        {
            using HttpResponseMessage response = await client.PostAsync(
                BaseUrl + "/MES000000/LoginCheck",
                content,
                cancellationToken);
            string body = await response.Content.ReadAsStringAsync(cancellationToken);
            return response.IsSuccessStatusCode &&
                   body.Contains("\"Result\":\"M\"", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static int ParseColumnIndex(string columnName)
    {
        string digits = new(columnName.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            ? index
            : 0;
    }

    private static string ReadText(JsonElement element, string propertyName)
    {
        if (!TryGetProperty(element, propertyName, out JsonElement value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return string.Empty;
        }

        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : value.GetRawText();
    }

    private static double? ReadNullableDouble(JsonElement element, string propertyName)
        => TryGetProperty(element, propertyName, out JsonElement value)
            ? ReadNullableDouble(value)
            : null;

    private static double? ReadNullableDouble(JsonElement value)
    {
        if (value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
            return null;
        if (value.ValueKind == JsonValueKind.Number)
            return value.TryGetDouble(out double numeric) ? numeric : null;
        if (value.ValueKind != JsonValueKind.String)
            return null;

        string text = value.GetString()?.Trim() ?? string.Empty;
        if (text.Length == 0)
            return null;
        return double.TryParse(
            text,
            NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.InvariantCulture,
            out double parsed)
            ? parsed
            : null;
    }

    private static bool TryGetProperty(
        JsonElement element,
        string propertyName,
        out JsonElement value)
    {
        if (element.ValueKind == JsonValueKind.Object &&
            element.TryGetProperty(propertyName, out value))
        {
            return true;
        }

        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (JsonProperty property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }
}

public sealed record IpgDefectParseResult(
    IReadOnlyList<IpgDefectRow> Rows,
    IReadOnlyList<IpgDefectColumn> Columns,
    int ColumnCount);

public sealed record IpgDefectColumn(
    int ColIndex,
    string ColumnName,
    string Header);

public sealed record IpgDefectKpiSnapshot(
    IReadOnlyList<IpgDefectKpiPeriod> Periods,
    double? AnnualAveragePpm);

public sealed record IpgDefectKpiPeriod(
    int ColIndex,
    string ColumnName,
    string Header,
    string Kind,
    double? AveragePpm);

public sealed class IpgDefectRow
{
    private readonly double?[] _columns = new double?[14];

    public string RawJson { get; set; } = string.Empty;
    public string Cptyp { get; set; } = string.Empty;
    public string CptypTx { get; set; } = string.Empty;
    public string FaccoTx { get; set; } = string.Empty;
    public string ModnoTx { get; set; } = string.Empty;
    public string ZlrtpTx { get; set; } = string.Empty;
    public string MatnrTx { get; set; } = string.Empty;
    public string Facco { get; set; } = string.Empty;
    public string Modno { get; set; } = string.Empty;
    public string Zsort { get; set; } = string.Empty;
    public string Zlrtp { get; set; } = string.Empty;
    public double? Avegr { get; set; }

    public double? Col0001 { get => GetCol(1); set => SetCol(1, value); }
    public double? Col0002 { get => GetCol(2); set => SetCol(2, value); }
    public double? Col0003 { get => GetCol(3); set => SetCol(3, value); }
    public double? Col0004 { get => GetCol(4); set => SetCol(4, value); }
    public double? Col0005 { get => GetCol(5); set => SetCol(5, value); }
    public double? Col0006 { get => GetCol(6); set => SetCol(6, value); }
    public double? Col0007 { get => GetCol(7); set => SetCol(7, value); }
    public double? Col0008 { get => GetCol(8); set => SetCol(8, value); }
    public double? Col0009 { get => GetCol(9); set => SetCol(9, value); }
    public double? Col0010 { get => GetCol(10); set => SetCol(10, value); }
    public double? Col0011 { get => GetCol(11); set => SetCol(11, value); }
    public double? Col0012 { get => GetCol(12); set => SetCol(12, value); }
    public double? Col0013 { get => GetCol(13); set => SetCol(13, value); }
    public double? Col0014 { get => GetCol(14); set => SetCol(14, value); }

    public double? GetCol(int index)
        => index is >= 1 and <= 14 ? _columns[index - 1] : null;

    public void SetCol(int index, double? value)
    {
        if (index is >= 1 and <= 14)
            _columns[index - 1] = value;
    }
}
