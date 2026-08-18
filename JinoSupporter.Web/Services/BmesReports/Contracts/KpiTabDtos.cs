namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public sealed class KpiTabDto
{
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public IReadOnlyList<KpiMetricDto> Metrics { get; init; } = [];
}

public sealed class KpiMetricDto
{
    public string Id { get; init; } = string.Empty;
    public string Name { get; init; } = string.Empty;
    public string Type { get; init; } = string.Empty;
    public double? BaselineValue { get; init; }
    public double? TargetValue { get; init; }
    public string Unit { get; init; } = "none";
    public IReadOnlyList<KpiLineDto> Lines { get; init; } = [];
}

public sealed class KpiLineDto
{
    public string Kind { get; init; } = string.Empty;
    public string Label { get; init; } = string.Empty;
    public string Unit { get; init; } = "none";
    public double? AnnualValue { get; init; }
    public Dictionary<string, double?> ValuesByPeriod { get; init; } = new(StringComparer.Ordinal);
}
