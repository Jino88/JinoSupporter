using System.IO;
using System.Text;
using System.Text.Json;

namespace InferenceDataAIService.Wpf;

internal sealed class AppPathSettings
{
    internal const string DefaultDataset = "InputDataFinish";

    public string ServiceDirectory { get; set; } = string.Empty;
    public string PythonExecutable { get; set; } = string.Empty;
    public string CodexExecutable { get; set; } = string.Empty;
    public string PrimaryDatabasePath { get; set; } = string.Empty;
    public string HistoryDatabasePath { get; set; } = string.Empty;
    public string OutputRootDirectory { get; set; } = string.Empty;
    public string BatchRootDirectory { get; set; } = string.Empty;
    public string RunLogDirectory { get; set; } = string.Empty;
    public string TemporaryDirectory { get; set; } = string.Empty;
    public string ExcelArchiveDirectory { get; set; } = string.Empty;
    public string IncrementalIngestDirectory { get; set; } = string.Empty;
    public string CorpusIngestDirectory { get; set; } = string.Empty;
    public string FormPreflightDirectory { get; set; } = string.Empty;
    public string EvidenceOutputDirectory { get; set; } = string.Empty;
    public string EvidenceDetailDirectory { get; set; } = string.Empty;
    public string RelatedStudiesDirectory { get; set; } = string.Empty;
    public string HumanReviewDirectory { get; set; } = string.Empty;
    public string AnalysisManifestDirectory { get; set; } = string.Empty;
    public string TableFirstRelevanceDirectory { get; set; } = string.Empty;
    public string BenchmarkResultPath { get; set; } = string.Empty;

    internal static AppPathSettings CreateDefaults(
        string serviceDirectory)
    {
        var service = Path.GetFullPath(serviceDirectory);
        var output = Path.Combine(service, "outputs");
        return CreateForOutputRoot(
            service,
            output,
            DefaultPythonExecutable(),
            DefaultCodexExecutable());
    }

    private static AppPathSettings CreateForOutputRoot(
        string service,
        string output,
        string pythonExecutable,
        string codexExecutable)
    {
        var universalGrid = Path.Combine(
            output,
            "universal-grid");
        return new AppPathSettings
        {
            ServiceDirectory = service,
            PythonExecutable = pythonExecutable,
            CodexExecutable = codexExecutable,
            PrimaryDatabasePath = Path.Combine(
                universalGrid,
                "InputDataFinish.sqlite"),
            HistoryDatabasePath = Path.Combine(
                output,
                "table-first-history",
                "history.sqlite"),
            OutputRootDirectory = output,
            BatchRootDirectory = Path.Combine(output, "batches"),
            RunLogDirectory = Path.Combine(output, "logs"),
            TemporaryDirectory = Path.Combine(output, "wpf-temp"),
            ExcelArchiveDirectory = Path.Combine(
                universalGrid,
                "ExcelFileArchive"),
            IncrementalIngestDirectory = Path.Combine(
                output,
                "incremental-ingest"),
            CorpusIngestDirectory = Path.Combine(
                output,
                "incremental-com-corpus"),
            FormPreflightDirectory = Path.Combine(
                output,
                "form-preflight"),
            EvidenceOutputDirectory = Path.Combine(
                output,
                "wpf-evidence"),
            EvidenceDetailDirectory = Path.Combine(
                output,
                "wpf-evidence-details"),
            RelatedStudiesDirectory = Path.Combine(
                output,
                "wpf-related-studies"),
            HumanReviewDirectory = Path.Combine(
                output,
                "wpf-human-review"),
            AnalysisManifestDirectory = Path.Combine(
                output,
                "analysis-manifests"),
            TableFirstRelevanceDirectory = Path.Combine(
                output,
                "table-first-relevance-answers"),
            BenchmarkResultPath = Path.Combine(
                output,
                "corpus-ingest",
                "full-989-v1",
                "benchmark-small-30-v23.result.json"),
        };
    }

    internal AppPathSettings Normalize(
        AppPathSettings defaults)
    {
        var service = FullDirectory(
            ServiceDirectory,
            defaults.ServiceDirectory,
            defaults.ServiceDirectory);
        var output = FullDirectory(
            OutputRootDirectory,
            defaults.OutputRootDirectory,
            service);
        return CreateForOutputRoot(
            service,
            output,
            Executable(
                PythonExecutable,
                defaults.PythonExecutable),
            Executable(
                CodexExecutable,
                defaults.CodexExecutable));
    }

    internal void Validate()
    {
        if (!Directory.Exists(ServiceDirectory))
            throw new DirectoryNotFoundException(
                $"서비스 폴더를 찾을 수 없습니다: {ServiceDirectory}");
        var cliPath = Path.Combine(
            ServiceDirectory,
            "inference_data_ai_cli.py");
        if (!File.Exists(cliPath))
            throw new FileNotFoundException(
                "서비스 폴더에 inference_data_ai_cli.py가 없습니다.",
                cliPath);
        ValidateExecutable(PythonExecutable, "Python");
        ValidateExecutable(CodexExecutable, "Codex");
    }

    private static string DefaultPythonExecutable()
    {
        var configured = Environment.GetEnvironmentVariable(
            "INFERENCE_DATA_AI_PYTHON");
        return string.IsNullOrWhiteSpace(configured)
            ? "python"
            : configured.Trim();
    }

    private static string DefaultCodexExecutable()
    {
        var configured = Environment.GetEnvironmentVariable(
            "INFERENCE_DATA_AI_CODEX");
        if (!string.IsNullOrWhiteSpace(configured))
            return configured.Trim();
        var npmShim = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.ApplicationData),
            "npm",
            "codex.cmd");
        return File.Exists(npmShim) ? npmShim : "codex";
    }

    private static string FullDirectory(
        string? value,
        string fallback,
        string baseDirectory) =>
        FullPath(value, fallback, baseDirectory);

    private static string FullPath(
        string? value,
        string fallback,
        string baseDirectory)
    {
        var selected = string.IsNullOrWhiteSpace(value)
            ? fallback
            : value.Trim();
        return Path.GetFullPath(
            Path.IsPathRooted(selected)
                ? selected
                : Path.Combine(baseDirectory, selected));
    }

    private static string Executable(
        string? value,
        string fallback)
    {
        var selected = string.IsNullOrWhiteSpace(value)
            ? fallback
            : value.Trim();
        return Path.IsPathRooted(selected)
            ? Path.GetFullPath(selected)
            : selected;
    }

    private static void ValidateExecutable(
        string executable,
        string label)
    {
        if (Path.IsPathRooted(executable)
            && !File.Exists(executable))
        {
            throw new FileNotFoundException(
                $"{label} 실행 파일을 찾을 수 없습니다.",
                executable);
        }
    }
}

internal static class AppPathSettingsStore
{
    private static readonly UTF8Encoding Utf8WithoutBom =
        new(encoderShouldEmitUTF8Identifier: false);

    internal static string SettingsFilePath => Path.Combine(
        Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData),
        "InferenceDataAIService",
        "settings.json");

    internal static AppPathSettings Load(
        string discoveredServiceDirectory)
    {
        var defaults = AppPathSettings.CreateDefaults(
            discoveredServiceDirectory);
        if (!File.Exists(SettingsFilePath))
            return defaults;
        try
        {
            var raw = JsonSerializer.Deserialize<AppPathSettings>(
                File.ReadAllText(SettingsFilePath, Encoding.UTF8));
            if (raw is null)
                return defaults;
            var configuredService = string.IsNullOrWhiteSpace(
                    raw.ServiceDirectory)
                ? defaults.ServiceDirectory
                : Path.GetFullPath(
                    Path.IsPathRooted(raw.ServiceDirectory)
                        ? raw.ServiceDirectory
                        : Path.Combine(
                            discoveredServiceDirectory,
                            raw.ServiceDirectory));
            return raw.Normalize(
                AppPathSettings.CreateDefaults(
                    configuredService));
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or JsonException
            or ArgumentException
            or NotSupportedException)
        {
            return defaults;
        }
    }

    internal static void Save(AppPathSettings settings)
    {
        var directory = Path.GetDirectoryName(SettingsFilePath)
            ?? throw new InvalidOperationException(
                "설정 저장 폴더를 확인할 수 없습니다.");
        Directory.CreateDirectory(directory);
        var temporary = SettingsFilePath
            + "."
            + Guid.NewGuid().ToString("N")
            + ".tmp";
        try
        {
            File.WriteAllText(
                temporary,
                JsonSerializer.Serialize(
                    settings,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,
                    }) + Environment.NewLine,
                Utf8WithoutBom);
            File.Move(
                temporary,
                SettingsFilePath,
                overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }
}

internal static class AppRuntimePaths
{
    private static AppPathSettings? _current;

    internal static AppPathSettings Current =>
        _current ?? throw new InvalidOperationException(
            "애플리케이션 경로 설정이 초기화되지 않았습니다.");

    internal static void Apply(AppPathSettings settings) =>
        _current = settings;
}

internal enum PathSettingKind
{
    Directory,
    File,
    Executable,
}

internal sealed class PathSettingRow(
    string key,
    string label,
    string value,
    PathSettingKind kind,
    string description)
{
    public string Key { get; } = key;
    public string Label { get; } = label;
    public string Value { get; set; } = value;
    public PathSettingKind Kind { get; } = kind;
    public string Description { get; } = description;
}
