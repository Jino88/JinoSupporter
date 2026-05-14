using System.Globalization;
using System.Text.Json.Serialization;

namespace BmesNgRateStandalone.Services;

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

public sealed class EvidenceComparison
{
    [JsonPropertyName("label")]      public string Label      { get; set; } = "";
    [JsonPropertyName("value")]      public string Value      { get; set; } = "";
    [JsonPropertyName("n")]          public int    N          { get; set; }
    [JsonPropertyName("isBaseline")] public bool   IsBaseline { get; set; }
    [JsonPropertyName("isBest")]     public bool   IsBest     { get; set; }
    [JsonPropertyName("isWorst")]    public bool   IsWorst    { get; set; }
}

public sealed class EvidenceRow
{
    [JsonPropertyName("metric")]         public string Metric         { get; set; } = "";
    // 2-arm 경로 (기존)
    [JsonPropertyName("baselineLabel")]  public string BaselineLabel  { get; set; } = "";
    [JsonPropertyName("baselineValue")]  public string BaselineValue  { get; set; } = "";
    [JsonPropertyName("variantLabel")]   public string VariantLabel   { get; set; } = "";
    [JsonPropertyName("variantValue")]   public string VariantValue   { get; set; } = "";
    [JsonPropertyName("deltaText")]      public string DeltaText      { get; set; } = "";
    [JsonPropertyName("deltaSign")]      public string DeltaSign      { get; set; } = "";  // up | down | no_change
    [JsonPropertyName("note")]           public string Note           { get; set; } = "";
    // 3+ arm 경로 (v7) — 비어있으면 위 2-arm 경로 사용
    [JsonPropertyName("comparisons")]    public List<EvidenceComparison>? Comparisons { get; set; }
    [JsonPropertyName("bestLabel")]      public string BestLabel      { get; set; } = "";
    [JsonPropertyName("worstLabel")]     public string WorstLabel     { get; set; } = "";
}

/// <summary>DOE 격자 실험 — Factor1 × Factor2 grid.</summary>
public sealed class DoeCell
{
    [JsonPropertyName("f1")]     public string F1     { get; set; } = "";   // factor1 level (e.g. "T=380")
    [JsonPropertyName("f2")]     public string F2     { get; set; } = "";   // factor2 level (e.g. "Tension=4")
    [JsonPropertyName("status")] public string Status { get; set; } = "";   // ok | ng | borderline | empty
    [JsonPropertyName("value")]  public string Value  { get; set; } = "";   // 측정값 또는 비고
}

public sealed class DoeGrid
{
    [JsonPropertyName("factor1Name")] public string       Factor1Name { get; set; } = "";   // e.g. "Temperature"
    [JsonPropertyName("factor2Name")] public string       Factor2Name { get; set; } = "";   // e.g. "Tension"
    [JsonPropertyName("factor1Levels")] public List<string> Factor1Levels { get; set; } = [];
    [JsonPropertyName("factor2Levels")] public List<string> Factor2Levels { get; set; } = [];
    [JsonPropertyName("cells")]       public List<DoeCell> Cells       { get; set; } = [];
}

/// <summary>시계열/주차별 trend.</summary>
public sealed class TrendPoint
{
    [JsonPropertyName("label")] public string Label { get; set; } = "";   // "Week 17", "Mar 2025"
    [JsonPropertyName("value")] public string Value { get; set; } = "";   // "8.3%"
    [JsonPropertyName("note")]  public string Note  { get; set; } = "";
}

public sealed class ActionItem
{
    [JsonPropertyName("priority")] public int    Priority { get; set; }
    [JsonPropertyName("kind")]     public string Kind     { get; set; } = "action";  // action | investigate | risk
    [JsonPropertyName("text")]     public string Text     { get; set; } = "";
}

public sealed class AnalysisContext
{
    [JsonPropertyName("process")]        public string Process        { get; set; } = "";
    [JsonPropertyName("stage")]          public string Stage          { get; set; } = "";
    [JsonPropertyName("baselineReason")] public string BaselineReason { get; set; } = "";
}

public sealed class NormalizeResult
{
    [JsonPropertyName("measurements")]        public List<NormalizedMeasurement> Measurements { get; set; } = [];
    [JsonPropertyName("tags")]                public List<string> Tags                { get; set; } = [];

    // ── v7 reportType — 카드 분기 ──
    // comparison_study | multi_arm | doe_factorial | reliability_validation
    // | trend_analysis | quality_log | intervention_test
    [JsonPropertyName("reportType")]          public string             ReportType { get; set; } = "";

    // ── v2 structured fields (verdict-first) ──
    [JsonPropertyName("verdict")]             public string             Verdict   { get; set; } = "";  // enum (v7: passed/failed 추가)
    [JsonPropertyName("headline")]            public string             Headline  { get; set; } = "";  // 1-line conclusion
    [JsonPropertyName("evidence")]            public List<EvidenceRow>  Evidence  { get; set; } = [];  // ≤4 rows
    [JsonPropertyName("actions")]             public List<ActionItem>   Actions   { get; set; } = [];  // ≤3 items
    [JsonPropertyName("context")]             public AnalysisContext?   Context   { get; set; }

    // ── v7 reportType-specific payloads (optional, set when reportType matches) ──
    [JsonPropertyName("doeGrid")]             public DoeGrid?           DoeGrid     { get; set; }
    [JsonPropertyName("trendPoints")]         public List<TrendPoint>?  TrendPoints { get; set; }

    // ── Legacy narrative fields (kept for backward-compat reading; empty in v2 output) ──
    [JsonPropertyName("summary")]             public string Summary             { get; set; } = "";
    [JsonPropertyName("keyFindings")]         public string KeyFindings         { get; set; } = "";
    [JsonPropertyName("purpose")]             public string Purpose             { get; set; } = "";
    [JsonPropertyName("testConditions")]      public string TestConditions      { get; set; } = "";
    [JsonPropertyName("rootCause")]           public string RootCause           { get; set; } = "";
    [JsonPropertyName("decision")]            public string Decision            { get; set; } = "";
    [JsonPropertyName("recommendedAction")]   public string RecommendedAction   { get; set; } = "";
}

// ── AI_EXCEL_PROC.md schema read DTOs (Batch CLI v6 output) ───────────────────
// Used by DataInferenceDbPage when the row was processed by the new CLI flow
// (writes AiDocuments / AiConclusions / AiTroubleshootingHints + ko/en/vi
// translations). Old DatasetSummary card is rendered for legacy rows.
public sealed class AiDocBundle
{
    public string DocumentId        { get; set; } = "";
    public string SourceDataset     { get; set; } = "";
    public string SourceFile        { get; set; } = "";
    public string Title             { get; set; } = "";
    public string Purpose           { get; set; } = "";
    public string PrimaryDefect     { get; set; } = "";
    public string ReportType        { get; set; } = "";
    public string ReportDate        { get; set; } = "";
    public double Confidence        { get; set; }
    public string DecisionRationale { get; set; } = "";   // from AiExtractionLogs

    public int ConditionsCount { get; set; }

    public List<string> Content        { get; set; } = [];
    public List<string> RelatedDefects { get; set; } = [];
    public List<string> Parts          { get; set; } = [];
    public List<string> Processes      { get; set; } = [];
    public List<string> Assumptions    { get; set; } = [];
    public List<string> Warnings       { get; set; } = [];

    public List<AiResultRow>             Results      { get; set; } = [];
    public List<AiNgBreakdownSummaryRow> NgBreakdowns { get; set; } = [];
    public List<AiConclusionRow> Conclusions { get; set; } = new();
    public List<AiHintRow>       Hints       { get; set; } = new();

    // Lang ('ko'|'en'|'vi') → translated narrative payload
    public Dictionary<string, AiDocTranslationRow>                          DocTranslations        { get; set; } = new();
    public Dictionary<string, Dictionary<string, AiConclusionTranslationRow>> ConclusionTranslations { get; set; } = new();
    public Dictionary<string, Dictionary<string, AiHintTranslationRow>>       HintTranslations       { get; set; } = new();
    public Dictionary<string, AiLogTranslationRow>                          LogTranslations        { get; set; } = new();
}

public sealed class AiConclusionRow
{
    public string ConclusionId             { get; set; } = "";
    public string Topic                    { get; set; } = "";
    public string StatementFromReport      { get; set; } = "";
    public string NormalizedInterpretation { get; set; } = "";
    public string SourceFile               { get; set; } = "";
    public string SheetName                { get; set; } = "";
}

public sealed class AiHintRow
{
    public string HintId           { get; set; } = "";
    public string DefectName       { get; set; } = "";
    public string CheckItem        { get; set; } = "";
    public string Reason           { get; set; } = "";
    public string EvidenceStrength { get; set; } = "";
    public string RelatedProcess   { get; set; } = "";
    public string RelatedPart      { get; set; } = "";
    public string SourceFile       { get; set; } = "";
    public string SheetName        { get; set; } = "";
}

public sealed class AiDocTranslationRow
{
    public string Title   { get; set; } = "";
    public string Purpose { get; set; } = "";
    public List<string> Content { get; set; } = [];
}
public sealed class AiConclusionTranslationRow
{
    public string Topic                    { get; set; } = "";
    public string StatementFromReport      { get; set; } = "";
    public string NormalizedInterpretation { get; set; } = "";
}
public sealed class AiHintTranslationRow
{
    public string CheckItem { get; set; } = "";
    public string Reason    { get; set; } = "";
}
public sealed class AiLogTranslationRow
{
    public string DecisionRationale { get; set; } = "";
    public List<string> Assumptions  { get; set; } = [];
    public List<string> Warnings     { get; set; } = [];
}

public sealed class AiResultRow
{
    public string  ResultId        { get; set; } = "";
    public string  ConditionId     { get; set; } = "";
    public string  MeasurementType { get; set; } = "";
    public string  ConditionGroup  { get; set; } = "";
    public string  ConditionProcess { get; set; } = "";
    public string  ChangedFactor   { get; set; } = "";
    public string  BeforeValue     { get; set; } = "";
    public string  AfterValue      { get; set; } = "";
    public string  ResultDate      { get; set; } = "";
    public string  Line            { get; set; } = "";
    public double? InputCount      { get; set; }
    public double? OkCount         { get; set; }
    public double? NgCount         { get; set; }
    public double? NgRateDecimal   { get; set; }
    public double? NgRatePercent   { get; set; }
    public string  MetricName      { get; set; } = "";
    public double? MetricValue     { get; set; }
    public string  Unit            { get; set; } = "";
    public string  Judgement       { get; set; } = "";
    public string  SourceFile      { get; set; } = "";
    public string  SheetName       { get; set; } = "";
    public string  SourceCellsJson { get; set; } = "";
}

public sealed class AiNgBreakdownSummaryRow
{
    public string  DefectName { get; set; } = "";
    public int     RowCount   { get; set; }
    public double  TotalCount { get; set; }
    public double? AvgRate    { get; set; }
}

public sealed class DatasetSummaryRecord
{
    public List<string> Tags { get; set; } = [];

    // v7 reportType
    public string                  ReportType { get; set; } = "";

    // v2 structured fields
    public string                  Verdict   { get; set; } = "";
    public string                  Headline  { get; set; } = "";
    public List<EvidenceRow>       Evidence  { get; set; } = [];
    public List<ActionItem>        Actions   { get; set; } = [];
    public AnalysisContext?        Context   { get; set; }

    // v7 reportType-specific payloads
    public DoeGrid?                DoeGrid     { get; set; }
    public List<TrendPoint>?       TrendPoints { get; set; }

    // Legacy narrative fields (still populated on pre-v2 rows)
    public string Summary           { get; set; } = "";
    public string KeyFindings       { get; set; } = "";
    public string Purpose           { get; set; } = "";
    public string TestConditions    { get; set; } = "";
    public string RootCause         { get; set; } = "";
    public string Decision          { get; set; } = "";
    public string RecommendedAction { get; set; } = "";

    // Translations: keyed by 2-letter lang code ("ko", "vi", …). The original
    // AI output is stored in the base columns above (treated as "en"); ko/vi
    // are filled by a follow-up translate call after Normalize. UI picks one.
    public Dictionary<string, DatasetSummaryTranslation> Translations { get; set; } = new();
}

public sealed class DatasetSummaryTranslation
{
    // v2 — translate-eligible fields only (numbers/labels stay verbatim)
    public string                  Headline { get; set; } = "";
    public List<ActionItem>        Actions  { get; set; } = [];
    public AnalysisContext?        Context  { get; set; }

    // Legacy
    public string Summary           { get; set; } = "";
    public string KeyFindings       { get; set; } = "";
    public string Purpose           { get; set; } = "";
    public string TestConditions    { get; set; } = "";
    public string RootCause         { get; set; } = "";
    public string Decision          { get; set; } = "";
    public string RecommendedAction { get; set; } = "";
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
