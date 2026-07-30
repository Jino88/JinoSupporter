namespace BmesNgRateStandalone.Services;

public static class CodexApiService
{
    public const string DefaultTranslateModel = "gpt-5.5";
}

public sealed record CurrentProblemApplyResult(
    int ReadRows,
    int MatchedRows,
    int SummaryRows,
    int ProductTypesFilled,
    int ReportDatesFilled,
    int MissingRows);

public sealed class CurrentProblemFirstPassRow
{
    public string DatasetName { get; set; } = "";
    public string FileNames { get; set; } = "";
    public string DbProductType { get; set; } = "";
    public string DbReportDate { get; set; } = "";
    public string AiModel { get; set; } = "";
    public string Model { get; set; } = "";
    public string ModelMappingSource { get; set; } = "";
    public string Date { get; set; } = "";
    public string PurposeCode { get; set; } = "";
    public string ReviewPurpose { get; set; } = "";
    public string Purpose { get; set; } = "";
    public List<string> TargetDefects { get; set; } = [];
    public List<string> ReviewItems { get; set; } = [];
    public List<string> Tags { get; set; } = [];
    public double Confidence { get; set; }
    public bool NeedsDetailedAnalysis { get; set; }
    public string EvidenceSummary { get; set; } = "";
    public List<string> EvidenceCells { get; set; } = [];
    public string Uncertainty { get; set; } = "";
}

public sealed record InputDataTestAnalysisParameters(
    string ReviewPurpose,
    IReadOnlyList<string> Tags,
    string Purpose,
    string PurposeCode,
    IReadOnlyList<string> TargetDefects,
    IReadOnlyList<string> ReviewItems,
    string Model,
    string Date,
    double? Confidence)
{
    public static InputDataTestAnalysisParameters Empty { get; } = new(
        "",
        Array.Empty<string>(),
        "",
        "",
        Array.Empty<string>(),
        Array.Empty<string>(),
        "",
        "",
        null);
}
