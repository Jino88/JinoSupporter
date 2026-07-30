using System.Collections.Generic;
using System.IO;

namespace InferenceDataAIService.Wpf;

public partial class App : System.Windows.Application
{
    private void Application_Startup(object sender, System.Windows.StartupEventArgs e)
    {
        var arguments = e.Args.Length > 0 ? e.Args : Environment.GetCommandLineArgs().Skip(1).ToArray();
        if (!StartupAnalysisOptions.TryParse(arguments, out var options, out var error))
        {
            System.Windows.MessageBox.Show(error, "시작 인자 오류", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            Shutdown(-1);
            return;
        }

        var window = new MainWindow(options);
        MainWindow = window;
        window.Show();
        window.ScheduleStartupAnalysis();
    }
}

internal sealed record StartupAnalysisOptions(IReadOnlyList<string> ExcelPaths, bool Force)
{
    private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx", ".xlsm", ".xls", ".xlsb",
    };

    public string? BatchFolder { get; init; }
    public string? BatchId { get; init; }
    public string? ResumeBatchId { get; init; }
    public string? NumericReviewBatchId { get; init; }
    public string? AiGroupAnalysisBatchId { get; init; }
    public string? BatchLimitArgument { get; init; }
    public int? BatchLimit { get; init; }
    public bool RetryFailed { get; init; }
    public bool RunAiGroupAnalysis { get; init; }

    public bool IsBatchFolderScan => !string.IsNullOrWhiteSpace(BatchFolder);
    public bool IsBatchResume => !string.IsNullOrWhiteSpace(ResumeBatchId);
    public bool IsBatchScan => IsBatchFolderScan || IsBatchResume;
    public bool IsNumericReview => !string.IsNullOrWhiteSpace(NumericReviewBatchId);
    public bool IsAiGroupAnalysis => !string.IsNullOrWhiteSpace(AiGroupAnalysisBatchId);
    public bool HasStartupWork => ExcelPaths.Count > 0 || IsBatchScan || IsNumericReview || IsAiGroupAnalysis;

    private const string Usage = "사용법:\n"
        + "InferenceDataAIService.Wpf.exe --analyze <Excel 전체 경로> [<Excel 전체 경로> ...] [--force]\n"
        + "또는 --analyze <Excel 전체 경로>를 반복할 수 있습니다.\n\n"
        + "InferenceDataAIService.Wpf.exe --batch-folder <폴더 전체 경로> [--batch-id ID] [--pilot N|--limit N]\n"
        + "  [--ai-group-analysis]\n"
        + "InferenceDataAIService.Wpf.exe --resume-batch <ID> [--limit N] [--retry-failed]\n"
        + "InferenceDataAIService.Wpf.exe --numeric-review-batch <ID>\n\n"
        + "InferenceDataAIService.Wpf.exe --ai-group-analysis-batch <ID>\n\n"
        + "배치 모드는 구조 사전 스캔만 실행합니다. COM, DB 적재, 분석 runner는 호출하지 않습니다.";

    public static bool TryParse(string[] arguments, out StartupAnalysisOptions options, out string error)
    {
        options = new StartupAnalysisOptions([], false);
        error = string.Empty;
        if (arguments.Length == 0) return true;

        var paths = new List<string>();
        var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var force = false;
        string? batchFolder = null;
        string? batchId = null;
        string? resumeBatchId = null;
        string? numericReviewBatchId = null;
        string? aiGroupAnalysisBatchId = null;
        string? batchLimitArgument = null;
        int? batchLimit = null;
        var retryFailed = false;
        var runAiGroupAnalysis = false;

        for (var index = 0; index < arguments.Length;)
        {
            var argument = arguments[index];
            if (string.Equals(argument, "--force", StringComparison.OrdinalIgnoreCase))
            {
                if (force)
                {
                    error = $"--force를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                force = true;
                index++;
                continue;
            }

            if (string.Equals(argument, "--analyze", StringComparison.OrdinalIgnoreCase))
            {
                index++;
                var firstPathIndex = index;
                while (index < arguments.Length && !arguments[index].StartsWith("--", StringComparison.Ordinal))
                {
                    if (!TryValidateExcelPath(arguments[index], out var fullPath, out error))
                    {
                        error = $"{error}\n\n{Usage}";
                        return false;
                    }
                    if (seenPaths.Add(fullPath)) paths.Add(fullPath);
                    index++;
                }
                if (index == firstPathIndex)
                {
                    error = $"--analyze 뒤에 하나 이상의 Excel 전체 경로가 필요합니다.\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--batch-folder", StringComparison.OrdinalIgnoreCase))
            {
                if (batchFolder is not null)
                {
                    error = $"--batch-folder를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                if (!TryReadOptionValue(arguments, ref index, "--batch-folder", out var input, out error) ||
                    !TryValidateBatchFolder(input, out batchFolder, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--resume-batch", StringComparison.OrdinalIgnoreCase))
            {
                if (resumeBatchId is not null)
                {
                    error = $"--resume-batch를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                if (!TryReadOptionValue(arguments, ref index, "--resume-batch", out var input, out error) ||
                    !TryValidateBatchId(input, out resumeBatchId, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--numeric-review-batch", StringComparison.OrdinalIgnoreCase))
            {
                if (numericReviewBatchId is not null)
                {
                    error = $"--numeric-review-batch를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                if (!TryReadOptionValue(arguments, ref index, "--numeric-review-batch", out var input, out error) ||
                    !TryValidateBatchId(input, out numericReviewBatchId, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--ai-group-analysis-batch", StringComparison.OrdinalIgnoreCase))
            {
                if (aiGroupAnalysisBatchId is not null || !TryReadOptionValue(arguments, ref index, "--ai-group-analysis-batch", out var input, out error) ||
                    !TryValidateBatchId(input, out aiGroupAnalysisBatchId, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--batch-id", StringComparison.OrdinalIgnoreCase))
            {
                if (batchId is not null)
                {
                    error = $"--batch-id를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                if (!TryReadOptionValue(arguments, ref index, "--batch-id", out var input, out error) ||
                    !TryValidateBatchId(input, out batchId, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                continue;
            }

            if (string.Equals(argument, "--pilot", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(argument, "--limit", StringComparison.OrdinalIgnoreCase))
            {
                if (batchLimit is not null)
                {
                    error = $"--pilot 또는 --limit는 한 번만 지정할 수 있습니다.\n\n{Usage}";
                    return false;
                }
                var optionName = argument.ToLowerInvariant();
                if (!TryReadOptionValue(arguments, ref index, optionName, out var input, out error) ||
                    !TryValidatePositiveInteger(input, optionName, out var parsedLimit, out error))
                {
                    error = $"{error}\n\n{Usage}";
                    return false;
                }
                batchLimitArgument = optionName;
                batchLimit = parsedLimit;
                continue;
            }

            if (string.Equals(argument, "--retry-failed", StringComparison.OrdinalIgnoreCase))
            {
                if (retryFailed)
                {
                    error = $"--retry-failed를 두 번 지정할 수 없습니다.\n\n{Usage}";
                    return false;
                }
                retryFailed = true;
                index++;
                continue;
            }

            if (string.Equals(argument, "--ai-group-analysis", StringComparison.OrdinalIgnoreCase))
            {
                runAiGroupAnalysis = true;
                index++;
                continue;
            }

            error = $"알 수 없는 시작 인자: {argument}\n\n{Usage}";
            return false;
        }

        var modeCount = (paths.Count > 0 ? 1 : 0) + (batchFolder is null ? 0 : 1) + (resumeBatchId is null ? 0 : 1) + (numericReviewBatchId is null ? 0 : 1) + (aiGroupAnalysisBatchId is null ? 0 : 1);
        if (modeCount > 1)
        {
            error = $"--analyze, --batch-folder, --resume-batch는 함께 사용할 수 없습니다.\n\n{Usage}";
            return false;
        }

        if (paths.Count > 0)
        {
            if (batchId is not null || batchLimit is not null || retryFailed)
            {
                error = $"배치 옵션은 --analyze와 함께 사용할 수 없습니다.\n\n{Usage}";
                return false;
            }
            options = new StartupAnalysisOptions(paths, force);
            return true;
        }

        if (batchFolder is not null)
        {
            if (force || retryFailed)
            {
                error = $"--force 및 --retry-failed는 --batch-folder와 함께 사용할 수 없습니다.\n\n{Usage}";
                return false;
            }
            options = new StartupAnalysisOptions([], false)
            {
                BatchFolder = batchFolder,
                BatchId = batchId,
                BatchLimitArgument = batchLimitArgument,
                BatchLimit = batchLimit,
                RunAiGroupAnalysis = runAiGroupAnalysis,
            };
            return true;
        }

        if (resumeBatchId is not null)
        {
            if (force || batchId is not null || string.Equals(batchLimitArgument, "--pilot", StringComparison.OrdinalIgnoreCase))
            {
                error = $"--force, --batch-id, --pilot는 --resume-batch와 함께 사용할 수 없습니다.\n\n{Usage}";
                return false;
            }
            options = new StartupAnalysisOptions([], false)
            {
                ResumeBatchId = resumeBatchId,
                BatchLimitArgument = batchLimitArgument,
                BatchLimit = batchLimit,
                RetryFailed = retryFailed,
            };
            return true;
        }

        if (numericReviewBatchId is not null)
        {
            if (force || batchId is not null || batchLimit is not null || retryFailed)
            {
                error = $"--force, --batch-id, --pilot, --limit, --retry-failed는 --numeric-review-batch와 함께 사용할 수 없습니다.\n\n{Usage}";
                return false;
            }
            options = new StartupAnalysisOptions([], false)
            {
                NumericReviewBatchId = numericReviewBatchId,
            };
            return true;
        }

        if (aiGroupAnalysisBatchId is not null)
        {
            if (force || batchId is not null || batchLimit is not null || retryFailed)
            {
                error = $"AI group analysis cannot be combined with other batch options.\n\n{Usage}";
                return false;
            }
            options = new StartupAnalysisOptions([], false) { AiGroupAnalysisBatchId = aiGroupAnalysisBatchId };
            return true;
        }

        if (force || batchId is not null || batchLimit is not null || retryFailed)
        {
            error = $"시작 옵션에는 --analyze, --batch-folder, 또는 --resume-batch가 필요합니다.\n\n{Usage}";
            return false;
        }

        options = new StartupAnalysisOptions([], false);
        return true;
    }

    private static bool TryReadOptionValue(string[] arguments, ref int index, string optionName, out string value, out string error)
    {
        value = string.Empty;
        error = string.Empty;
        if (index + 1 >= arguments.Length || arguments[index + 1].StartsWith("--", StringComparison.Ordinal))
        {
            error = $"{optionName} 뒤에 값이 필요합니다.";
            return false;
        }
        value = arguments[index + 1];
        index += 2;
        return true;
    }

    private static bool TryValidateBatchFolder(string input, out string fullPath, out string error)
    {
        fullPath = string.Empty;
        error = string.Empty;
        if (!Path.IsPathFullyQualified(input))
        {
            error = $"배치 폴더 경로는 전체 경로여야 합니다: {input}";
            return false;
        }
        try
        {
            fullPath = Path.GetFullPath(input);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error = $"유효한 배치 폴더 경로가 아닙니다: {input}";
            return false;
        }
        if (!Directory.Exists(fullPath))
        {
            error = $"배치 폴더를 찾을 수 없습니다: {input}";
            return false;
        }
        return true;
    }

    private static bool TryValidateBatchId(string input, out string batchId, out string error)
    {
        batchId = string.Empty;
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(input) || input is "." or ".." || input.Length > 96 ||
            input.Any(character => !(char.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '.')))
        {
            error = "배치 ID는 영문/숫자/.-_만 사용해 1~96자로 지정해야 합니다.";
            return false;
        }
        batchId = input;
        return true;
    }

    private static bool TryValidatePositiveInteger(string input, string optionName, out int value, out string error)
    {
        value = 0;
        error = string.Empty;
        if (!int.TryParse(input, out value) || value < 1)
        {
            error = $"{optionName}에는 1 이상의 정수를 지정해야 합니다: {input}";
            return false;
        }
        return true;
    }

    private static bool TryValidateExcelPath(string input, out string fullPath, out string error)
    {
        fullPath = string.Empty;
        error = string.Empty;
        if (!Path.IsPathFullyQualified(input))
        {
            error = $"Excel 경로는 전체 경로여야 합니다: {input}";
            return false;
        }
        try
        {
            fullPath = Path.GetFullPath(input);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            error = $"유효한 Excel 경로가 아닙니다: {input}";
            return false;
        }
        if (!ExcelExtensions.Contains(Path.GetExtension(fullPath)))
        {
            error = $"지원하지 않는 Excel 확장자입니다: {input}";
            return false;
        }
        if (Path.GetFileName(fullPath).StartsWith("~$", StringComparison.Ordinal))
        {
            error = $"Excel 임시 파일은 분석할 수 없습니다: {input}";
            return false;
        }
        if (!File.Exists(fullPath))
        {
            error = $"Excel 파일을 찾을 수 없습니다: {input}";
            return false;
        }
        return true;
    }
}
