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
    [JsonIgnore] public long Id { get; set; }
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
    [JsonPropertyName("measurements")]        public List<NormalizedMeasurement> Measurements { get; set; } = [];
    [JsonPropertyName("summary")]             public string       Summary             { get; set; } = "";
    [JsonPropertyName("keyFindings")]         public string       KeyFindings         { get; set; } = "";
    [JsonPropertyName("tags")]                public List<string> Tags                { get; set; } = [];

    // Structured context fields for AI Ask — extract from report's common sections
    [JsonPropertyName("purpose")]             public string       Purpose             { get; set; } = "";
    [JsonPropertyName("testConditions")]      public string       TestConditions      { get; set; } = "";
    [JsonPropertyName("rootCause")]           public string       RootCause           { get; set; } = "";
    [JsonPropertyName("decision")]            public string       Decision            { get; set; } = "";
    [JsonPropertyName("recommendedAction")]   public string       RecommendedAction   { get; set; } = "";
}

public sealed class DatasetSummaryRecord
{
    public string       Summary           { get; set; } = "";
    public string       KeyFindings       { get; set; } = "";
    public List<string> Tags              { get; set; } = [];

    // Extended structured fields
    public string       Purpose           { get; set; } = "";
    public string       TestConditions    { get; set; } = "";
    public string       RootCause         { get; set; } = "";
    public string       Decision          { get; set; } = "";
    public string       RecommendedAction { get; set; } = "";
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
    string CreatedAt,
    bool   BatchExcluded,
    string BatchedAt = "");

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

public sealed class ModelGroupRecord
{
    public long                 Id           { get; set; }
    public string               Name         { get; set; } = string.Empty;   // 대그룹 이름
    public string               ProductGroup { get; set; } = "ETC";          // SPK / UNIT / MODULE / TWS / ETC
    public int                  SortOrder    { get; set; }
    public List<MidGroupRecord> MidGroups    { get; set; } = new();
}

public sealed class MidGroupRecord
{
    public string                Material  { get; set; } = string.Empty;   // 중그룹 (MAKTX)
    public List<SubGroupRecord>  SubGroups { get; set; } = new();          // 세그룹 (recursive tree)

    /// <summary>Flattened view across the entire sub-group subtree. Read-only shim so
    /// existing report code (By Model / By Group) keeps working without caring about
    /// sub-group structure.</summary>
    public IReadOnlyList<string> LineShifts =>
        SubGroups.SelectMany(s => s.AllLineShifts).ToList();
}

public sealed class SubGroupRecord
{
    public string               Name       { get; set; } = string.Empty;   // 비어 있으면 "기본" 버킷
    public List<string>         LineShifts { get; set; } = new();
    public List<SubGroupRecord> SubGroups  { get; set; } = new();          // 중첩 서브그룹 (재귀)

    /// <summary>Depth-first flattening of LineShifts across this node and all descendants.</summary>
    public IEnumerable<string> AllLineShifts =>
        LineShifts.Concat(SubGroups.SelectMany(s => s.AllLineShifts));
}

public sealed class BmesMaterial
{
    public string Matnr      { get; set; } = string.Empty;
    public string Maktx      { get; set; } = string.Empty;
    public string Meins      { get; set; } = string.Empty;
    public string Injtp      { get; set; } = string.Empty;
    public string Mtype      { get; set; } = string.Empty;
    public string Btype      { get; set; } = string.Empty;
    public string MngCode    { get; set; } = string.Empty;
    public string ModNameB   { get; set; } = string.Empty;
    public string LotQt      { get; set; } = string.Empty;
    public string Bunch      { get; set; } = string.Empty;
    public string NgTar      { get; set; } = string.Empty;
    public string McLv1Tx    { get; set; } = string.Empty;
    public string McLv2Tx    { get; set; } = string.Empty;
    public string McLv3Tx    { get; set; } = string.Empty;
    public string McLv4Tx    { get; set; } = string.Empty;
    public string McLv5Tx    { get; set; } = string.Empty;
    public string McLv6Tx    { get; set; } = string.Empty;
    public string Ernam      { get; set; } = string.Empty;
    public string Erdat      { get; set; } = string.Empty;
    public string Grcod      { get; set; } = string.Empty;
    public string Grnam      { get; set; } = string.Empty;
    public string MfPhi      { get; set; } = string.Empty;
    public string FetchedAt  { get; set; } = string.Empty;
}
