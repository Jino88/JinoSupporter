namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public sealed class CauseMonthlyTabDto
{
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public IReadOnlyList<CauseRowDto> Rows { get; init; } = [];
    public IReadOnlyList<CauseModelMonthlyRowDto> ModelMonthlyRows { get; init; } = [];
}

public sealed class CauseRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string Model { get; init; } = string.Empty;
    public string? Type { get; init; }
    public string? Process { get; init; }
    public string? NgName { get; init; }
    public int? Number { get; init; }
    public string? Cause { get; init; }
    public double? ShareRatio { get; init; }
    public bool IsSubtotal { get; init; }
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> WeightedPpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class CauseModelMonthlyRowDto
{
    public string Model { get; init; } = string.Empty;
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}
