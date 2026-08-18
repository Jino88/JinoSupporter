namespace JinoSupporter.Web.Services.BmesReports.Contracts;

public enum ReportState
{
    Complete,
    Partial,
    Failed,
}

public sealed record ReportIssueDto(
    string Code,
    string Message,
    string Source,
    bool Retryable);

public sealed class ReportStatusDto
{
    public ReportState State { get; init; } = ReportState.Complete;
    public string Code { get; init; } = "ok";
    public string? Message { get; init; }
    public IReadOnlyList<ReportIssueDto> Warnings { get; init; } = [];
    public IReadOnlyList<ReportIssueDto> Errors { get; init; } = [];

    public static ReportStatusDto Complete() => new();

    public static ReportStatusDto Partial(IReadOnlyList<ReportIssueDto> warnings, string? message = null) => new()
    {
        State = ReportState.Partial,
        Code = warnings.FirstOrDefault()?.Code ?? "partial",
        Message = message,
        Warnings = warnings,
    };

    public static ReportStatusDto Failed(ReportIssueDto error) => new()
    {
        State = ReportState.Failed,
        Code = error.Code,
        Message = error.Message,
        Errors = [error],
    };
}

public sealed class ReportPeriodDto
{
    public string Key { get; init; } = string.Empty;
    public string Kind { get; init; } = string.Empty;
    public string Header { get; init; } = string.Empty;
    public int SortOrder { get; init; }
    public DateOnly? StartDate { get; init; }
    public DateOnly? EndDateExclusive { get; init; }
    public int? SourceIndex { get; init; }
    public string? SourceCode { get; init; }
    public string? SourcePDate { get; init; }
}

public sealed class ReportTabEnvelope<T>
{
    public ReportStatusDto Status { get; init; } = ReportStatusDto.Complete();
    public T? Data { get; init; }

    public static ReportTabEnvelope<T> Complete(T data) => new() { Data = data };

    public static ReportTabEnvelope<T> Partial(T data, IReadOnlyList<ReportIssueDto> warnings) => new()
    {
        Status = ReportStatusDto.Partial(warnings),
        Data = data,
    };
}

public static class BmesReportContract
{
    public const string ContractId = "jinosupporter.bmes-report";
    public const string SchemaVersion = "1.0.0";
    public const string CalculationVersion = "legacy-2026-08-18";
}
