namespace BmesNgRateStandalone.Services.BmesReports.Contracts;

public sealed class DailyTabDto
{
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public IReadOnlyList<DailySummaryRowDto> SummaryRows { get; init; } = [];
    public IReadOnlyList<DailyModelSectionDto> ModelSections { get; init; } = [];
    public string? ReferenceDatePeriodKey { get; init; }
}

public sealed class DailySummaryRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public int Depth { get; init; }
    public string Level { get; init; } = string.Empty;
    public string? GroupName { get; init; }
    public string? ModelName { get; init; }
    public string Display { get; init; } = string.Empty;
    public string? ProcessType { get; init; }
    public Dictionary<string, double?> InputByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> NgByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class DailyModelSectionDto
{
    public string Id { get; init; } = string.Empty;
    public string GroupName { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public int LineShiftCount { get; init; }
    public IReadOnlyList<DailyReasonRowDto> ReasonRows { get; init; } = [];
}

public sealed class DailyReasonRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public string Reason { get; init; } = string.Empty;
    public int? Rank { get; init; }
    public bool IsTotal { get; init; }
    public string? ProcessType { get; init; }
    public string? ProcessName { get; init; }
    public string? NgName { get; init; }
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}
