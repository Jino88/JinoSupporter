using System.Globalization;
using System.Text;
using System.Text.Json;

namespace PortalGNLAB.Web.Services;

/// <summary>
/// Pure conversions between a PLATFORM_API Tool answer ({ result, message, resultSets:[{columns, rows}] })
/// and the row shapes the BMES screens used to return. No network, so the unit tests feed it saved answers.
/// Column names of the Tools are the BMES screen's own column names, so a result set becomes the same
/// "contents" objects the screen's SearchList gave and the screen parsers (F-Cost, Core parts) run unchanged.
/// </summary>
public static class BmesToolRows
{
    public static BmesApiResult ParseExecuteResponse(string json)
    {
        using JsonDocument doc = JsonDocument.Parse(json);
        JsonElement root = doc.RootElement;
        var sets = new List<BmesApiResultSet>();
        if (root.TryGetProperty("resultSets", out var resultSets) && resultSets.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement set in resultSets.EnumerateArray())
            {
                List<string> columns = set.TryGetProperty("columns", out var c) && c.ValueKind == JsonValueKind.Array
                    ? c.EnumerateArray().Select(x => x.GetString() ?? "").ToList()
                    : [];
                List<JsonElement[]> rows = set.TryGetProperty("rows", out var r) && r.ValueKind == JsonValueKind.Array
                    ? r.EnumerateArray().Select(row => row.EnumerateArray().Select(v => v.Clone()).ToArray()).ToList()
                    : [];
                sets.Add(new BmesApiResultSet(columns, rows));
            }
        }

        return new BmesApiResult(Text(root, "toolName"), Text(root, "result"), Text(root, "message"), sets);
    }

    /// <summary>First result set that has all the given columns (a Tool's result sets differ per view).</summary>
    public static BmesApiResultSet? FindSet(BmesApiResult result, params string[] columns) =>
        result.ResultSets
            .Where(s => columns.All(c => s.Column(c) >= 0))
            .OrderByDescending(s => s.Rows.Count)
            .FirstOrDefault();

    /// <summary>Cell as the text the BMES screens gave: null → "", true/false → "Y"/"N" is NOT applied here.</summary>
    public static string CellText(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null or JsonValueKind.Undefined => string.Empty,
        JsonValueKind.String => value.GetString() ?? string.Empty,
        _ => value.ToString(),
    };

    /// <summary>One column→text dictionary per row (keys case-insensitive), optionally renaming columns.</summary>
    public static List<Dictionary<string, string>> ToRowDictionaries(
        BmesApiResultSet set, IReadOnlyDictionary<string, string>? rename = null, params string[] skipColumns)
    {
        var skip = new HashSet<string>(skipColumns, StringComparer.OrdinalIgnoreCase);
        var rows = new List<Dictionary<string, string>>(set.Rows.Count);
        foreach (JsonElement[] source in set.Rows)
        {
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < set.Columns.Count && i < source.Length; i++)
            {
                string name = set.Columns[i];
                if (skip.Contains(name)) continue;
                if (rename is not null && rename.TryGetValue(name, out string? mapped)) name = mapped;
                row[name] = CellText(source[i]);
            }
            rows.Add(row);
        }
        return rows;
    }

    /// <summary>
    /// The screen's SearchList wrapper: { "result": true, "data": { "contents": [ {col: value} ],
    /// "BottomGridColumnList": [ {ZCOLN, ZPVTT, ZPVTT_TX, PDATE} ] } }. Values keep their JSON type.
    /// </summary>
    public static string ToSearchListJson(BmesApiResultSet grid, BmesApiResultSet? periodColumns = null)
    {
        using var stream = new MemoryStream();
        using (var w = new Utf8JsonWriter(stream))
        {
            w.WriteStartObject();
            w.WriteBoolean("result", true);
            w.WriteStartObject("data");
            w.WritePropertyName("contents");
            WriteObjects(w, grid);
            if (periodColumns is not null)
            {
                w.WritePropertyName("BottomGridColumnList");
                WriteObjects(w, periodColumns);
            }
            w.WriteEndObject();
            w.WriteEndObject();
        }
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    private static void WriteObjects(Utf8JsonWriter w, BmesApiResultSet set)
    {
        w.WriteStartArray();
        foreach (JsonElement[] row in set.Rows)
        {
            w.WriteStartObject();
            for (int i = 0; i < set.Columns.Count && i < row.Length; i++)
            {
                w.WritePropertyName(set.Columns[i]);
                row[i].WriteTo(w);
            }
            w.WriteEndObject();
        }
        w.WriteEndArray();
    }

    /// <summary>Generic rows for the LPA screens: column order as returned, every value as text
    /// (null → "", true/false → "Y"/"N"), lookups by column name are case-insensitive.</summary>
    public static (List<string> Columns, List<IReadOnlyDictionary<string, string>> Rows) ToGenericRows(BmesApiResultSet set)
    {
        var rows = new List<IReadOnlyDictionary<string, string>>(set.Rows.Count);
        foreach (JsonElement[] source in set.Rows)
        {
            var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < set.Columns.Count && i < source.Length; i++)
            {
                JsonElement v = source[i];
                row[set.Columns[i]] = v.ValueKind switch
                {
                    JsonValueKind.True => "Y",
                    JsonValueKind.False => "N",
                    _ => CellText(v),
                };
            }
            rows.Add(row);
        }
        return (set.Columns.ToList(), rows);
    }

    /// <summary>Consecutive windows of at most <paramref name="maxMonths"/> calendar months covering [start, end]
    /// (2026-07-28 → 2026-10-27 is one 3-month window).</summary>
    public static List<(DateTime Start, DateTime End)> SplitPeriod(DateTime start, DateTime end, int maxMonths = BmesApi.MaxPeriodMonths)
    {
        var windows = new List<(DateTime, DateTime)>();
        for (DateTime s = start.Date; s <= end.Date; )
        {
            DateTime e = s.AddMonths(maxMonths).AddDays(-1);
            if (e > end.Date) e = end.Date;
            windows.Add((s, e));
            s = e.AddDays(1);
        }
        return windows;
    }

    /// <summary>"yyyyMMdd" or "yyyy-MM-dd" → "yyyy-MM-dd" (the Tools take ISO dates); empty stays empty.</summary>
    public static string ToIsoDate(string value)
    {
        string text = value.Trim();
        if (text.Length == 0) return text;
        string[] formats = ["yyyyMMdd", "yyyy-MM-dd"];
        return DateTime.TryParseExact(text, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime d)
            ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
            : text;
    }

    private static string Text(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";
}
