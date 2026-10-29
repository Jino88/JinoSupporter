using System.Globalization;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

public static partial class BmesGateway
{
    /// <summary>The API call for MES072410: AI_TOOL_00001 view G (Core Part F-Cost grid, daily = 7 days + 4 weeks +
    /// 3 months), plant GN like the screen call.</summary>
    public static Dictionary<string, object> CorePartsParameters(DateTime queryDate) => new()
    {
        ["ZTYPE"] = "G",
        ["SDATE"] = queryDate.ToString("yyyy-MM-dd"),
        ["ZGUBN"] = "D",
        ["FACCO"] = "GN",
    };

    /// <summary>
    /// MES072410 core-parts F-COST from PLATFORM_API (AI_TOOL_00001, view G).
    /// Input: <paramref name="queryDate"/> (SDATE, daily mode ZGUBN=D, plant GN).
    /// Output: rows (INAMT / FCOST / FRATE triplets with Col0001..Col0014 values), the period column
    /// metadata and the column count. Throws <see cref="BmesApiException"/> on an API failure.
    /// </summary>
    public static async Task<FCostCorePartsParseResult> FetchCoreParts(
        BmesWebSession session,
        DateTime queryDate,
        CancellationToken cancellationToken = default)
    {
        BmesApiResult result = await BmesApi.ExecuteAsync(
            RequireApiCredentials(session), "AI_TOOL_00001", CorePartsParameters(queryDate),
            cancellationToken: cancellationToken);
        return CorePartsFromResult(result);
    }

    /// <summary>Tool answer (view G) → the same parse the screen's SearchList went through: the grid is the result set
    /// with CPTYP_TX / COL0001, the period columns the set with ZCOLN / PDATE. An empty answer gives no rows.</summary>
    public static FCostCorePartsParseResult CorePartsFromResult(BmesApiResult result)
    {
        BmesApiResultSet? grid = BmesToolRows.FindSet(result, "CPTYP_TX", "COL0001", "ZTYPE");
        if (grid is null) return new FCostCorePartsParseResult([], [], 0);
        return ParseCorePartsSearchList(BmesToolRows.ToSearchListJson(grid, BmesToolRows.FindSet(result, "ZCOLN", "PDATE")));
    }

    public static FCostCorePartsParseResult ParseCorePartsSearchList(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        if (!CorePartsProp(root, "data", out JsonElement data) ||
            data.ValueKind != JsonValueKind.Object)
        {
            throw new JsonException("MES072410 response does not contain data.");
        }

        if (!CorePartsProp(data, "contents", out JsonElement contents) ||
            contents.ValueKind != JsonValueKind.Array)
        {
            throw new JsonException("MES072410 response does not contain data.contents.");
        }

        var rows = new List<FCostCorePartsRow>(contents.GetArrayLength());
        int columnCount = 0;
        foreach (JsonElement item in contents.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.Object)
                continue;

            var row = new FCostCorePartsRow
            {
                RawJson = item.GetRawText(),
                FaccoTx = CorePartsText(item, "FACCO_TX"),
                CptypTx = CorePartsText(item, "CPTYP_TX"),
                ZtypeTx = CorePartsText(item, "ZTYPE_TX"),
                Ztote = CorePartsText(item, "ZTOTE"),
                Zsort01 = CorePartsText(item, "ZSORT_01"),
                Zsort02 = CorePartsText(item, "ZSORT_02"),
                Zsort03 = CorePartsText(item, "ZSORT_03"),
                Zsort04 = CorePartsText(item, "ZSORT_04"),
                Cptyp = CorePartsText(item, "CPTYP"),
                Tsort = CorePartsText(item, "TSORT"),
                Zsort = CorePartsText(item, "ZSORT"),
                Facco = CorePartsText(item, "FACCO"),
                Ztype = CorePartsText(item, "ZTYPE"),
                Chk = CorePartsText(item, "CHK"),
            };

            for (int index = 1; index <= 14; index++)
            {
                string propertyName = $"COL{index:D4}";
                if (CorePartsProp(item, propertyName, out JsonElement value))
                {
                    columnCount = Math.Max(columnCount, index);
                    row.SetCol(index, CorePartsNumber(value));
                }
            }

            rows.Add(row);
        }

        ForwardFillTriplets(rows);
        List<FCostCorePartsColumn> columns = ParseTrustedColumnMetadata(data);
        if (columns.Count > 0)
            columnCount = Math.Max(columnCount, columns.Max(column => column.ColIndex));

        return new FCostCorePartsParseResult(rows, columns, columnCount);
    }

    private static List<FCostCorePartsColumn> ParseTrustedColumnMetadata(JsonElement data)
    {
        var result = new List<FCostCorePartsColumn>();
        if (!CorePartsProp(data, "BottomGridColumnList", out JsonElement metadata) ||
            metadata.ValueKind != JsonValueKind.Array)
        {
            return result;
        }

        foreach (JsonElement item in metadata.EnumerateArray())
        {
            string columnName = CorePartsText(item, "ZCOLN");
            int index = CorePartsColumnIndex(columnName);
            if (index is < 1 or > 14)
                continue;

            string code = CorePartsText(item, "ZPVTT");
            string header = CorePartsText(item, "ZPVTT_TX");
            string pdate = CorePartsText(item, "PDATE");
            if (string.IsNullOrWhiteSpace(code) &&
                string.IsNullOrWhiteSpace(header) &&
                string.IsNullOrWhiteSpace(pdate))
            {
                continue;
            }

            result.Add(new FCostCorePartsColumn(
                index,
                code,
                header,
                pdate,
                InferPeriodKind(code, pdate)));
        }

        return result
            .GroupBy(column => column.ColIndex)
            .Select(group => group.First())
            .OrderBy(column => column.ColIndex)
            .ToList();
    }

    private static string InferPeriodKind(string code, string pdate)
    {
        string candidate = string.IsNullOrWhiteSpace(code) ? pdate : code;
        if (candidate.Contains('-', StringComparison.Ordinal))
            return "Week";
        string digits = new(candidate.Where(char.IsDigit).ToArray());
        return digits.Length switch
        {
            8 => "Day",
            6 => "Month",
            _ => "Unknown",
        };
    }

    private static int CorePartsColumnIndex(string columnName)
    {
        string digits = new(columnName.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, NumberStyles.None, CultureInfo.InvariantCulture, out int index)
            ? index
            : 0;
    }

    private static void ForwardFillTriplets(IReadOnlyList<FCostCorePartsRow> rows)
    {
        FCostCorePartsRow? anchor = null;
        foreach (FCostCorePartsRow row in rows)
        {
            if (string.Equals(row.Ztype, "INAMT", StringComparison.OrdinalIgnoreCase))
            {
                anchor = row;
                continue;
            }

            bool continuation =
                string.Equals(row.Ztype, "FCOST", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(row.Ztype, "FRATE", StringComparison.OrdinalIgnoreCase);
            if (!continuation)
            {
                anchor = null;
                continue;
            }
            if (anchor is null)
                continue;

            if (string.IsNullOrWhiteSpace(row.FaccoTx)) row.FaccoTx = anchor.FaccoTx;
            if (string.IsNullOrWhiteSpace(row.CptypTx)) row.CptypTx = anchor.CptypTx;
            if (string.IsNullOrWhiteSpace(row.Facco)) row.Facco = anchor.Facco;
            if (string.IsNullOrWhiteSpace(row.Cptyp)) row.Cptyp = anchor.Cptyp;
        }
    }

    private static string CorePartsText(JsonElement element, string propertyName)
    {
        if (!CorePartsProp(element, propertyName, out JsonElement value) ||
            value.ValueKind is JsonValueKind.Null or JsonValueKind.Undefined)
        {
            return string.Empty;
        }

        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : value.GetRawText();
    }

    private static double? CorePartsNumber(JsonElement value)
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

    private static bool CorePartsProp(
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
