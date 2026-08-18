namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public sealed class WeeklyTabDto
{
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public IReadOnlyList<WeeklyTargetRowDto> TargetRows { get; init; } = [];
    public IReadOnlyList<WeeklySummaryRowDto> SummaryRows { get; init; } = [];
    public IReadOnlyList<WeeklyTrendSeriesDto> TrendSeries { get; init; } = [];
    public IReadOnlyList<WeeklyDefectRowDto> TopDefects { get; init; } = [];
    public string? SortReferencePeriodKey { get; init; }
}

public sealed class WeeklyTargetRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public string Display { get; init; } = string.Empty;
    public string RowKind { get; init; } = string.Empty;
    public int Depth { get; init; }
    public bool IsLineShift { get; init; }
    public double? BaselinePpm { get; init; }
    public double? TargetPpm { get; init; }
    public double? AchievementPercent { get; init; }
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class WeeklySummaryRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public int Depth { get; init; }
    public string Level { get; init; } = string.Empty;
    public string? GroupName { get; init; }
    public string? ModelName { get; init; }
    public string? ProcessType { get; init; }
    public string Display { get; init; } = string.Empty;
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class WeeklyTrendSeriesDto
{
    public string Id { get; init; } = string.Empty;
    public string Label { get; init; } = string.Empty;
    public string? GroupName { get; init; }
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class WeeklyDefectRowDto
{
    public int Rank { get; init; }
    public string? LineShift { get; init; }
    public string ProcessType { get; init; } = string.Empty;
    public string ProcessName { get; init; } = string.Empty;
    public string NgName { get; init; } = string.Empty;
    public Dictionary<string, double?> PpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}
