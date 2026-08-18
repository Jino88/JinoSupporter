using System.Text.Json;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public sealed class BmesReportDocumentDto
{
    public string ContractId { get; init; } = BmesReportContract.ContractId;
    public string SchemaVersion { get; init; } = BmesReportContract.SchemaVersion;
    public string CalculationVersion { get; init; } = BmesReportContract.CalculationVersion;
    public DateTimeOffset GeneratedAtUtc { get; init; } = DateTimeOffset.UtcNow;
    public BmesReportRequestDto Request { get; init; } = new();
    public BmesReportViewerDefaultsDto ViewerDefaults { get; init; } = new();
    public ReportStatusDto Status { get; init; } = ReportStatusDto.Complete();
    public BmesReportTabsDto Tabs { get; init; } = new();
}

public sealed class BmesReportRequestDto
{
    public DateOnly StartDate { get; init; }
    public DateOnly EndDate { get; init; }
    public string TimeZoneId { get; init; } = "Asia/Bangkok";
    public IReadOnlyList<ReportSelectionGroupDto> Groups { get; init; } = [];
}

public sealed class ReportSelectionGroupDto
{
    public long Id { get; init; }
    public string Name { get; init; } = string.Empty;
    public IReadOnlyList<ReportSelectionMidDto> MidGroups { get; init; } = [];
}

public sealed class ReportSelectionMidDto
{
    public string Material { get; init; } = string.Empty;
    public IReadOnlyList<string> LineShifts { get; init; } = [];
}

public sealed class BmesReportViewerDefaultsDto
{
    public string DefaultTab { get; init; } = "daily";
    public double MinimumPpm { get; init; } = 500d;
    public int DateColumnLimit { get; init; } = 7;
    public int WeekColumnLimit { get; init; } = 4;
    public int MonthColumnLimit { get; init; } = 3;
    public string FcostCurrency { get; init; } = "VND";
}

public sealed class BmesReportTabsDto
{
    public ReportTabEnvelope<DailyTabDto> Daily { get; init; } = new();
    public ReportTabEnvelope<WeeklyTabDto> Weekly { get; init; } = new();
    public ReportTabEnvelope<KpiTabDto> Kpi { get; init; } = new();

    [JsonPropertyName("cause-monthly")]
    public ReportTabEnvelope<CauseMonthlyTabDto> CauseMonthly { get; init; } = new();

    public ReportTabEnvelope<FCostTabDto> Fcost { get; init; } = new();

    [JsonPropertyName("fcost-all")]
    public ReportTabEnvelope<FCostFollowerTabDto> FcostAll { get; init; } = new();

    [JsonPropertyName("fcost-weekly")]
    public ReportTabEnvelope<FCostFollowerTabDto> FcostWeekly { get; init; } = new();

    [JsonPropertyName("fcost-weekly-all")]
    public ReportTabEnvelope<FCostFollowerTabDto> FcostWeeklyAll { get; init; } = new();
}

public static class BmesReportJson
{
    public static JsonSerializerOptions SerializerOptions { get; } = CreateOptions();

    public static string Serialize(BmesReportDocumentDto document) =>
        JsonSerializer.Serialize(document, SerializerOptions);

    public static byte[] SerializeToUtf8Bytes(BmesReportDocumentDto document) =>
        JsonSerializer.SerializeToUtf8Bytes(document, SerializerOptions);

    private static JsonSerializerOptions CreateOptions()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DictionaryKeyPolicy = null,
            WriteIndented = false,
            DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
        return options;
    }
}
