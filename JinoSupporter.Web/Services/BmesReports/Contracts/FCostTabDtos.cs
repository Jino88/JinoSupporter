namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public sealed class FCostTabDto
{
    public FCostViewDto View { get; init; } = new();
    public FCostDatasetDto Dataset { get; init; } = new();
}

public sealed class FCostFollowerTabDto
{
    public FCostViewDto View { get; init; } = new();
}

public sealed class FCostViewDto
{
    public string Mode { get; init; } = "regular";
    public bool AllPeriods { get; init; }
    public string? SourceTab { get; init; }
}

public sealed class FCostDatasetDto
{
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public FCostTotalsDto Totals { get; init; } = new();
    public IReadOnlyList<FCostTrendRowDto> TrendRows { get; init; } = [];
    public IReadOnlyList<FCostHierarchyRowDto> HierarchyRows { get; init; } = [];
    public IReadOnlyList<FCostMaterialRowDto> Materials { get; init; } = [];
    public IReadOnlyList<FCostMaterialRowDto> UnmappedMaterials { get; init; } = [];
    public FCostRawBreakdownDto? RawBreakdown { get; init; }
    public TargetDefectRateDto TargetDefectRate { get; init; } = new();
}

public sealed class FCostTotalsDto
{
    public Dictionary<string, double?> InputQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> FcostUsdByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> RatePercentByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class FCostTrendRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public int Depth { get; init; }
    public string GroupName { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string NgGroupKey { get; init; } = string.Empty;
    public Dictionary<string, double?> InputQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> FcostUsdByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> NgPpmByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> FcostSharePercentByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class FCostHierarchyRowDto
{
    public string RowId { get; init; } = string.Empty;
    public string? ParentRowId { get; init; }
    public int Depth { get; init; }
    public string GroupName { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string NgGroupKey { get; init; } = string.Empty;
    public string Display { get; init; } = string.Empty;
    public int MatchedMaterialCount { get; init; }
    public Dictionary<string, double?> InputQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> FcostUsdByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> SourceRatePercentByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class FCostMaterialRowDto
{
    public string DisplayName { get; init; } = string.Empty;
    public string? ProductGroup { get; init; }
    public string? ModelNo { get; init; }
    public string? Material { get; init; }
    public string? Verid { get; init; }
    public Dictionary<string, double?> InputQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> FcostUsdByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, double?> SourceRatePercentByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class FCostRawBreakdownDto
{
    public string SourceTable { get; init; } = string.Empty;
    public string NameSource { get; init; } = string.Empty;
    public string? WarningMessage { get; init; }
    public IReadOnlyList<ReportPeriodDto> Periods { get; init; } = [];
    public IReadOnlyList<FCostExchangeRateDto> ExchangeRates { get; init; } = [];
    public IReadOnlyList<FCostRawMaterialRowDto> Rows { get; init; } = [];
}

public sealed class FCostExchangeRateDto
{
    public string PeriodKey { get; init; } = string.Empty;
    public DateOnly StandardDate { get; init; }
    public decimal? KrwPerUsd { get; init; }
    public decimal? KrwPerVnd { get; init; }
}

public sealed class FCostRawMaterialRowDto
{
    public string GroupName { get; init; } = string.Empty;
    public string ModelName { get; init; } = string.Empty;
    public string MaterialCode { get; init; } = string.Empty;
    public string MaterialName { get; init; } = string.Empty;
    public Dictionary<string, decimal?> FcostVndByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, decimal?> EquivalentQtyByPeriod { get; init; } = new(StringComparer.Ordinal);
    public Dictionary<string, FCostRawPriceDto?> PriceByPeriod { get; init; } = new(StringComparer.Ordinal);
    public decimal TotalFcostVnd { get; init; }
    public long SourceRows { get; init; }
}

public sealed class FCostRawPriceDto
{
    public decimal? UnitPrice { get; init; }
    public string? Currency { get; init; }
    public string? PriceUnit { get; init; }
    public decimal? UnitPriceVnd { get; init; }
    public bool IsMixed { get; init; }
}

public sealed class TargetDefectRateDto
{
    public double DefaultBaselineRatePercent { get; init; } = 6.12;
    public IReadOnlyList<double> DefaultTargetRatePercent { get; init; } = [2.0, 0.7];
    public IReadOnlyList<string> ActionItems { get; init; } = [];
    public IReadOnlyList<TargetDefectRateRowDto> Rows { get; init; } = [];
    public Dictionary<string, double?> TotalActualPpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}

public sealed class TargetDefectRateRowDto
{
    public string ModelName { get; init; } = string.Empty;
    public double? TargetPpm { get; init; }
    public double? AchievementPercent { get; init; }
    public Dictionary<string, double?> ActualPpmByPeriod { get; init; } = new(StringComparer.Ordinal);
}
