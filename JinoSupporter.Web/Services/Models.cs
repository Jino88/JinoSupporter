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

public sealed class NormalizedMeasurement
{
    [JsonPropertyName("productType")]    public string ProductType    { get; set; } = "";
    [JsonPropertyName("testDate")]       public string TestDate       { get; set; } = "";
    [JsonPropertyName("line")]           public string Line           { get; set; } = "";
    [JsonPropertyName("checkType")]      public string CheckType      { get; set; } = "";
    [JsonPropertyName("variable")]       public string Variable       { get; set; } = "";
    [JsonPropertyName("variableDetail")] public string VariableDetail { get; set; } = "";
    [JsonPropertyName("variableGroup")]  public string VariableGroup  { get; set; } = "";
    [JsonPropertyName("intervention")]   public string Intervention   { get; set; } = "";
    [JsonPropertyName("inputQty")]       public int    InputQty       { get; set; }
    [JsonPropertyName("okQty")]          public int    OkQty          { get; set; }
    [JsonPropertyName("ngTotal")]        public int    NgTotal        { get; set; }
    [JsonPropertyName("ngRate")]         public double NgRate         { get; set; }
    [JsonPropertyName("defectCategory")] public string DefectCategory { get; set; } = "";
    [JsonPropertyName("defectType")]     public string DefectType     { get; set; } = "";
    [JsonPropertyName("defectCount")]    public int    DefectCount    { get; set; }
}

public sealed class NormalizeResult
{
    [JsonPropertyName("measurements")] public List<NormalizedMeasurement> Measurements { get; set; } = [];
    [JsonPropertyName("summary")]      public string       Summary     { get; set; } = "";
    [JsonPropertyName("keyFindings")]  public string       KeyFindings { get; set; } = "";
    [JsonPropertyName("tags")]         public List<string> Tags        { get; set; } = [];
}

public sealed class DatasetSummaryRecord
{
    public string       Summary     { get; set; } = "";
    public string       KeyFindings { get; set; } = "";
    public List<string> Tags        { get; set; } = [];
}

public sealed record RawFileInfo(
    long   Id,
    string FileName,
    string MediaType,
    long   FileSize,
    string CreatedAt);

public sealed record AskAiHistoryRecord(
    long   Id,
    string Question,
    string ProductTypeFilter,
    string Overall,
    string PerDatasetJson,
    string CreatedAt);

public sealed record RawReportInfo(
    long   Id,
    string DatasetName,
    string ProductType,
    string ReportDate,
    int    ImageCount,
    int    MeasurementCount,
    string CreatedAt);

public sealed record ImprovementRow(
    string  DatasetName,
    string  ProductType,
    string  TestDate,
    string  Line,
    string  CheckType,
    string  VariableDetail,
    string  DefectCategory,
    string  DefectType,
    double? NormalNgRate,
    double? TestNgRate,
    double? ImprovementPct,
    int     NormalInputQty,
    int     TestInputQty,
    string  Intervention);
