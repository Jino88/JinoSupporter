using System.Globalization;
using System.Text;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    /// <summary>
    /// MES050032 IPG defect rate, raw response.
    /// Input: <paramref name="queryDate"/> and <paramref name="periodCode"/> (W/WEEK or M/MONTH), plant GN,
    /// POSTed as a JSON body (see <see cref="BuildIpgSearchListPayload"/>).
    /// Output: the response JSON text (header + ChartRec); feed it to <see cref="ParseIpgSearchList"/>.
    /// Throws on HTTP failure or an invalid period code.
    /// </summary>
    public static async Task<string> FetchIpgDefectJson(
        BmesWebSession session,
        DateTime queryDate,
        string periodCode,
        CancellationToken cancellationToken = default)
    {
        string normalizedPeriodCode = NormalizeIpgPeriodCode(periodCode);
        string payload = BuildIpgSearchListPayload(queryDate, normalizedPeriodCode);
        if (await WebLoginErrorAsync(session, null, cancellationToken) is { } loginError)
            throw new InvalidOperationException(loginError);
        using var content = new StringContent(payload, Encoding.UTF8, "application/json");
        using HttpResponseMessage response = await session.Client.PostAsync(
            BmesWebSession.BaseUrl + "/MES050032/SearchList",
            content,
            cancellationToken);
        string json = await response.Content.ReadAsStringAsync(cancellationToken);
        response.EnsureSuccessStatusCode();
        return json;
    }

    /// <summary>MES050032 request body. Input: query date + period code (W or M). Output: JSON text.</summary>
    public static string BuildIpgSearchListPayload(DateTime queryDate, string periodCode)
    {
        string normalizedPeriodCode = NormalizeIpgPeriodCode(periodCode);
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

    /// <summary>
    /// MES050032 response parser. Input: response JSON text. Output: rows (one per ChartRec
    /// entry, Col0001..Col0014 values plus AVEGR), header columns and the column count.
    /// Throws JsonException when the server reports an error or ChartRec is missing.
    /// </summary>
    public static IpgDefectParseResult ParseIpgSearchList(string json)
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

    /// <summary>Normalizes W/WEEK/M/MONTH to "W" or "M"; throws ArgumentException otherwise.</summary>
    public static string NormalizeIpgPeriodCode(string periodCode) =>
        (periodCode ?? string.Empty).Trim().ToUpperInvariant() switch
        {
            "W" or "WEEK" => "W",
            "M" or "MONTH" => "M",
            _ => throw new ArgumentException(
                "MES050032 period must be W (week) or M (month).",
                nameof(periodCode)),
        };

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
