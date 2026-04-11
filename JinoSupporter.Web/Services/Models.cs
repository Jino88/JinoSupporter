using System.Globalization;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services;

public sealed class ColumnDef
{
    [JsonPropertyName("field")] public string Field { get; set; } = string.Empty;
    [JsonPropertyName("label")] public string Label { get; set; } = string.Empty;
}

public sealed class DataTableInfo
{
    public long            Id          { get; init; }
    public string          DatasetName { get; init; } = string.Empty;
    public string          TableName   { get; init; } = string.Empty;
    public List<ColumnDef> Columns     { get; init; } = [];
    public string          CreatedAt   { get; init; } = string.Empty;
    public int             RowCount    { get; init; }

    public string DisplayLabel => $"{TableName}  ({RowCount:N0} rows)";

    public string CreatedAtLocal
    {
        get
        {
            if (DateTime.TryParse(CreatedAt, null, DateTimeStyles.RoundtripKind, out DateTime dt))
                return dt.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
            return CreatedAt;
        }
    }
}

public sealed class ExtractedTable
{
    [JsonPropertyName("tableName")] public string TableName { get; set; } = string.Empty;
    [JsonPropertyName("columns")]   public List<ColumnDef> Columns { get; set; } = [];
    [JsonPropertyName("rows")]      public List<Dictionary<string, string>> Rows { get; set; } = [];
}
