using System.Text.Json;
using InferenceDataAIService.Wpf;

var options = AuditOptions.Parse(args);
var serviceDirectory = Path.GetFullPath(options.ServiceDirectory);
var pathSettings = AppPathSettingsStore.Load(serviceDirectory);
var databasePath = Path.GetFullPath(
    options.DatabasePath
    ?? pathSettings.PrimaryDatabasePath);
var client = new WorkbookComparisonClient(pathSettings);

IReadOnlyList<WorkbookComparisonSummary> summaries;
if (options.Scope == AuditScope.Representative30)
{
    summaries = await client.ListAsync(databasePath);
}
else
{
    var journalPath = Path.GetFullPath(
        options.JournalPath
        ?? Path.Combine(
            serviceDirectory,
            "outputs",
            "corpus-ingest",
            "full-989-v1",
            "corpus-journal.json"));
    var completedSources = ReadCompletedSources(journalPath);
    summaries = await client.ListSourcesAsync(
        databasePath,
        completedSources);
}

var issues = new List<AuditIssue>();
var observationCount = 0;
var matchedCount = 0;
var unavailableCount = 0;

foreach (var summary in summaries)
{
    if (!summary.IsAvailable)
    {
        unavailableCount++;
        issues.Add(new AuditIssue(
            summary.BenchmarkNumber,
            summary.FileName,
            summary.PublicAnalysisId,
            "DB 분석 없음",
            string.Empty,
            string.Empty,
            string.Empty,
            string.Empty));
        continue;
    }

    var document = await client.LoadAsync(
        databasePath,
        summary.PublicAnalysisId);
    foreach (var observation in document.Observations)
    {
        observationCount++;
        if (string.Equals(
                observation.ComparisonStatus,
                "일치",
                StringComparison.Ordinal))
        {
            matchedCount++;
            continue;
        }
        issues.Add(new AuditIssue(
            summary.BenchmarkNumber,
            summary.FileName,
            summary.PublicAnalysisId,
            observation.ProblemDescription,
            observation.ExcelActualValue,
            observation.DatabaseValue,
            observation.StudyTitle,
            observation.OutcomeLabel));
    }
}

var report = new AuditReport(
    options.Scope == AuditScope.Representative30
        ? "representative-30"
        : "completed-corpus",
    summaries.Count,
    unavailableCount,
    observationCount,
    matchedCount,
    issues.Count,
    issues
        .Select(issue => issue.FileNumber)
        .Distinct()
        .Count(),
    issues);

if (options.Json)
{
    Console.WriteLine(JsonSerializer.Serialize(
        report,
        new JsonSerializerOptions
        {
            WriteIndented = true,
        }));
    return;
}

Console.WriteLine(
    $"범위={report.Scope} 파일={report.FileCount} "
    + $"DB없음={report.UnavailableFileCount} "
    + $"값={report.ObservationCount} "
    + $"일치={report.MatchedCount} "
    + $"문제={report.IssueCount} "
    + $"문제파일={report.FilesWithIssues}");

if (issues.Count == 0)
{
    Console.WriteLine("문제 없음");
    return;
}

Console.WriteLine("문제 목록");
foreach (var issue in issues)
{
    Console.WriteLine(
        $"[{issue.FileNumber}] {issue.FileName}");
    Console.WriteLine(
        $"  문제: {issue.Problem}");
    Console.WriteLine(
        $"  Excel 실제값: {issue.ExcelActualValue}");
    Console.WriteLine(
        $"  DB 인식값: {issue.DatabaseValue}");
}

static IReadOnlyList<string> ReadCompletedSources(
    string journalPath)
{
    if (!File.Exists(journalPath))
        throw new FileNotFoundException(
            "corpus journal을 찾을 수 없습니다.",
            journalPath);

    using var document = JsonDocument.Parse(
        File.ReadAllText(journalPath));
    if (!document.RootElement.TryGetProperty(
            "records",
            out var records)
        || records.ValueKind != JsonValueKind.Array)
        throw new InvalidDataException(
            "corpus journal records가 없습니다.");

    var result = new List<string>();
    foreach (var record in records.EnumerateArray())
    {
        if (!record.TryGetProperty(
                "status",
                out var status)
            || !string.Equals(
                status.GetString(),
                "COMPLETED",
                StringComparison.Ordinal))
            continue;
        if (!record.TryGetProperty(
                "sourcePath",
                out var sourcePath))
            continue;
        var value = sourcePath.GetString();
        if (!string.IsNullOrWhiteSpace(value))
            result.Add(value);
    }
    return result;
}

internal enum AuditScope
{
    CompletedCorpus,
    Representative30,
}

internal sealed record AuditOptions(
    string ServiceDirectory,
    string? DatabasePath,
    string? JournalPath,
    AuditScope Scope,
    bool Json)
{
    internal static AuditOptions Parse(
        IReadOnlyList<string> args)
    {
        var serviceDirectory =
            Directory.GetCurrentDirectory();
        string? databasePath = null;
        string? journalPath = null;
        var scope = AuditScope.CompletedCorpus;
        var json = false;

        for (var index = 0; index < args.Count; index++)
        {
            var argument = args[index];
            switch (argument)
            {
                case "--service-dir":
                    serviceDirectory = ReadValue(
                        args,
                        ref index,
                        argument);
                    break;
                case "--db":
                    databasePath = ReadValue(
                        args,
                        ref index,
                        argument);
                    break;
                case "--journal":
                    journalPath = ReadValue(
                        args,
                        ref index,
                        argument);
                    break;
                case "--scope":
                    var value = ReadValue(
                        args,
                        ref index,
                        argument);
                    scope = value switch
                    {
                        "completed" =>
                            AuditScope.CompletedCorpus,
                        "representative-30" =>
                            AuditScope.Representative30,
                        _ => throw new ArgumentException(
                            "--scope는 completed 또는 "
                            + "representative-30이어야 합니다."),
                    };
                    break;
                case "--json":
                    json = true;
                    break;
                default:
                    throw new ArgumentException(
                        $"알 수 없는 인자: {argument}");
            }
        }

        return new AuditOptions(
            serviceDirectory,
            databasePath,
            journalPath,
            scope,
            json);
    }

    private static string ReadValue(
        IReadOnlyList<string> args,
        ref int index,
        string option)
    {
        index++;
        if (index >= args.Count)
            throw new ArgumentException(
                $"{option} 값이 필요합니다.");
        return args[index];
    }
}

internal sealed record AuditIssue(
    int FileNumber,
    string FileName,
    string AnalysisId,
    string Problem,
    string ExcelActualValue,
    string DatabaseValue,
    string Study,
    string Outcome);

internal sealed record AuditReport(
    string Scope,
    int FileCount,
    int UnavailableFileCount,
    int ObservationCount,
    int MatchedCount,
    int IssueCount,
    int FilesWithIssues,
    IReadOnlyList<AuditIssue> Issues);
