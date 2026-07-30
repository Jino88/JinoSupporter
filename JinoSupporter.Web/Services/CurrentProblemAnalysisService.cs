using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace JinoSupporter.Web.Services;

public sealed record CurrentProblemStatus(
    string DataRoot,
    string SampleDir,
    bool HasSampleDir,
    int IndexedRows,
    int TermCount,
    int PromptRuleCount,
    int ClassificationDone,
    int ClassificationTotal,
    string LatestResultHtmlPath,
    string DemoReportHtmlPath,
    string LatestAnalysisJsonPath,
    string SearchHtmlPath,
    string BatchControlHtmlPath,
    string TermGlossaryHtmlPath,
    string UpdatedAt);

public sealed record CurrentProblemRunRequest(
    string Problem,
    string Model,
    int Top,
    int CandidateLimit,
    int MeasureReportLimit,
    string Effort,
    string MeasureEffort,
    string HtmlEffort,
    int TimeoutSeconds,
    bool SkipMeasurements);

public sealed record CurrentProblemProcessResult(
    bool Success,
    int ExitCode,
    string Output,
    string Error,
    string JsonPath,
    string HtmlPath);

public sealed record CurrentProblemApplyResult(
    int ReadRows,
    int MatchedRows,
    int SummaryRows,
    int ProductTypesFilled,
    int ReportDatesFilled,
    int MissingRows);

public sealed class CurrentProblemFirstPassRow
{
    public string DatasetName { get; set; } = "";
    public string FileNames { get; set; } = "";
    public string DbProductType { get; set; } = "";
    public string DbReportDate { get; set; } = "";
    public string AiModel { get; set; } = "";
    public string Model { get; set; } = "";
    public string ModelMappingSource { get; set; } = "";
    public string Date { get; set; } = "";
    public string PurposeCode { get; set; } = "";
    public string ReviewPurpose { get; set; } = "";
    public string Purpose { get; set; } = "";
    public List<string> TargetDefects { get; set; } = [];
    public List<string> ReviewItems { get; set; } = [];
    public List<string> Tags { get; set; } = [];
    public double Confidence { get; set; }
    public bool NeedsDetailedAnalysis { get; set; }
    public string EvidenceSummary { get; set; } = "";
    public List<string> EvidenceCells { get; set; } = [];
    public string Uncertainty { get; set; } = "";
}

public sealed record CurrentProblemIndexRow(
    int No,
    string DatasetName,
    string Model,
    string Purpose,
    string PurposeDetail,
    IReadOnlyList<string> TargetDefects,
    IReadOnlyList<string> ReviewItems,
    IReadOnlyList<string> Tags,
    double Confidence);

public enum CurrentProblemArtifact
{
    DemoReportHtml,
    ResultHtml,
    SearchHtml,
    BatchControlHtml,
    TermGlossaryHtml,
    LatestAnalysisJson,
}

public sealed class CurrentProblemAnalysisService
{
    private readonly IWebHostEnvironment _env;
    private readonly string _dataRoot;
    private readonly SemaphoreSlim _processLock = new(1, 1);

    public CurrentProblemAnalysisService(IWebHostEnvironment env)
    {
        _env = env;
        _dataRoot = ResolveDataRoot(env);
    }

    public string DataRoot => _dataRoot;
    public string SampleDir => Path.Combine(DataRoot, "sample_ready");
    public bool HasFirstPassIndex => File.Exists(Path.Combine(SampleDir, "demo_index.json"));

    public bool SourceDirectoryHasFirstPassIndex(string sourceDirectory)
    {
        string source = Path.GetFullPath((sourceDirectory ?? "").Trim());
        if (!Directory.Exists(source))
            return false;

        return File.Exists(Path.Combine(source, "demo_index.json"))
               || File.Exists(Path.Combine(source, "sample_ready", "demo_index.json"));
    }

    public CurrentProblemStatus GetStatus()
    {
        string latestJson = LatestFile(Path.Combine(SampleDir, "ai_current_problem"), "current_problem_analysis_*.json");
        string demoReportHtml = Path.Combine(SampleDir, "demo_report.html");
        string resultHtml = Path.Combine(SampleDir, "current_problem_ai_analysis.html");
        string searchHtml = Path.Combine(SampleDir, "current_problem_search.html");
        string controlHtml = Path.Combine(SampleDir, "ai_batch_control.html");
        string glossaryHtml = Path.Combine(SampleDir, "ai_term_glossary.html");

        (int done, int total, string updatedAt) = ReadClassificationSummary();

        return new CurrentProblemStatus(
            DataRoot,
            SampleDir,
            Directory.Exists(SampleDir),
            CountJsonArray(Path.Combine(SampleDir, "demo_index.json")),
            CountMarkdownHeadings(Path.Combine(SampleDir, "ai_term_guidance.md")),
            CountMarkdownHeadings(Path.Combine(SampleDir, "prompt_update_requests.md")),
            done,
            total,
            File.Exists(resultHtml) ? resultHtml : "",
            File.Exists(demoReportHtml) ? demoReportHtml : "",
            latestJson,
            File.Exists(searchHtml) ? searchHtml : "",
            File.Exists(controlHtml) ? controlHtml : "",
            File.Exists(glossaryHtml) ? glossaryHtml : "",
            updatedAt);
    }

    public string ReadArtifact(CurrentProblemArtifact artifact)
    {
        string path = artifact switch
        {
            CurrentProblemArtifact.DemoReportHtml => Path.Combine(SampleDir, "demo_report.html"),
            CurrentProblemArtifact.ResultHtml => Path.Combine(SampleDir, "current_problem_ai_analysis.html"),
            CurrentProblemArtifact.SearchHtml => Path.Combine(SampleDir, "current_problem_search.html"),
            CurrentProblemArtifact.BatchControlHtml => Path.Combine(SampleDir, "ai_batch_control.html"),
            CurrentProblemArtifact.TermGlossaryHtml => Path.Combine(SampleDir, "ai_term_glossary.html"),
            CurrentProblemArtifact.LatestAnalysisJson => LatestFile(Path.Combine(SampleDir, "ai_current_problem"), "current_problem_analysis_*.json"),
            _ => "",
        };

        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return "";
        return File.ReadAllText(path, Encoding.UTF8);
    }

    public IReadOnlyList<CurrentProblemIndexRow> GetDemoIndexRows()
    {
        string path = Path.Combine(SampleDir, "demo_index.json");
        if (!File.Exists(path)) return [];

        try
        {
            using JsonDocument doc = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            if (doc.RootElement.ValueKind != JsonValueKind.Array) return [];

            var rows = new List<CurrentProblemIndexRow>();
            int no = 0;
            foreach (JsonElement item in doc.RootElement.EnumerateArray())
            {
                no++;
                string purpose = JsonString(item, "reviewPurpose");
                string purposeDetail = JsonString(item, "purpose");
                rows.Add(new CurrentProblemIndexRow(
                    no,
                    JsonString(item, "datasetName"),
                    JsonString(item, "model"),
                    string.IsNullOrWhiteSpace(purpose) ? purposeDetail : purpose,
                    purposeDetail,
                    JsonStringArray(item, "targetDefects"),
                    JsonStringArray(item, "reviewItems"),
                    JsonStringArray(item, "tags"),
                    JsonDouble(item, "confidence")));
            }

            return rows;
        }
        catch
        {
            return [];
        }
    }

    public async Task ImportFromDirectoryAsync(string sourceDirectory, CancellationToken ct = default)
    {
        string source = Path.GetFullPath((sourceDirectory ?? "").Trim());
        if (!Directory.Exists(source))
            throw new DirectoryNotFoundException(source);

        string destination = Path.GetFullPath(DataRoot);
        if (PathEquals(source, destination) || IsSubPathOf(destination, source))
            throw new InvalidOperationException("Import source cannot be the web current-problem data folder or its parent.");

        Directory.CreateDirectory(destination);

        string sourceName = Path.GetFileName(source.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        string target = string.Equals(sourceName, "sample_ready", StringComparison.OrdinalIgnoreCase)
            ? SampleDir
            : destination;

        await Task.Run(() => CopyDirectory(source, target, ct), ct);
    }

    public async Task<CurrentProblemApplyResult> ApplyFirstPassIndexToDbAsync(WebRepository repo, CancellationToken ct = default)
    {
        string path = Path.Combine(SampleDir, "demo_index.json");
        if (!File.Exists(path))
            throw new FileNotFoundException($"First-pass index not found: {path}", path);

        return await Task.Run(() =>
        {
            ct.ThrowIfCancellationRequested();
            string json = File.ReadAllText(path, Encoding.UTF8);
            List<CurrentProblemFirstPassRow> rows = JsonSerializer.Deserialize<List<CurrentProblemFirstPassRow>>(
                    json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? [];
            ct.ThrowIfCancellationRequested();
            return repo.ApplyCurrentProblemFirstPassRows(rows);
        }, ct);
    }

    public async Task<CurrentProblemProcessResult> RebuildDemoIndexAsync(int timeoutSeconds = 240, CancellationToken ct = default)
    {
        Directory.CreateDirectory(SampleDir);
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string script = Path.Combine(repoRoot, "prepare_ai_review_demo.py");
        string results = Path.Combine(DataRoot, "classification_results.jsonl");
        if (!File.Exists(results))
            throw new FileNotFoundException($"Classification results not found: {results}", results);

        var args = new List<string>
        {
            "--results", results,
            "--out-dir", SampleDir,
            "--max-items", "0",
        };

        string modelMap = Path.Combine(SampleDir, "model_mapping_conditions.csv");
        if (File.Exists(modelMap))
        {
            args.Add("--model-map");
            args.Add(modelMap);
        }

        var output = new StringBuilder();
        var error = new StringBuilder();

        CurrentProblemProcessResult demo = await RunPythonAsync(repoRoot, script, args, timeoutSeconds, ct);
        output.AppendLine(demo.Output);
        error.AppendLine(demo.Error);
        if (!demo.Success)
            return demo with { Output = output.ToString(), Error = error.ToString() };

        CurrentProblemProcessResult staticResult = await RebuildStaticArtifactsAsync(timeoutSeconds, ct);
        output.AppendLine(staticResult.Output);
        error.AppendLine(staticResult.Error);

        return staticResult with
        {
            Output = output.ToString(),
            Error = error.ToString(),
            HtmlPath = Path.Combine(SampleDir, "demo_report.html"),
        };
    }

    public async Task<CurrentProblemProcessResult> RebuildStaticArtifactsAsync(int timeoutSeconds = 180, CancellationToken ct = default)
    {
        Directory.CreateDirectory(SampleDir);
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string searchScript = Path.Combine(repoRoot, "create_current_problem_search_html.py");
        string controlScript = Path.Combine(repoRoot, "create_ai_batch_control_html.py");

        var output = new StringBuilder();
        var error = new StringBuilder();

        CurrentProblemProcessResult search = await RunPythonAsync(
            repoRoot,
            searchScript,
            ["--sample-dir", SampleDir],
            timeoutSeconds,
            ct);
        output.AppendLine(search.Output);
        error.AppendLine(search.Error);
        if (!search.Success)
            return search with { Output = output.ToString(), Error = error.ToString() };

        CurrentProblemProcessResult control = await RunPythonAsync(
            repoRoot,
            controlScript,
            ["--sample-dir", SampleDir, "--batch-dir", DataRoot],
            timeoutSeconds,
            ct);
        output.AppendLine(control.Output);
        error.AppendLine(control.Error);

        return control with
        {
            Output = output.ToString(),
            Error = error.ToString(),
            HtmlPath = Path.Combine(SampleDir, "current_problem_search.html"),
        };
    }

    public async Task<CurrentProblemProcessResult> RunFirstPassClassificationAsync(
        string dbPath,
        string effort,
        int limit,
        int timeoutSeconds,
        CancellationToken ct = default)
    {
        Directory.CreateDirectory(DataRoot);
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string script = Path.Combine(repoRoot, "ai_first_pass_classify.py");

        List<string> args =
        [
            "--db", dbPath,
            "--out-dir", DataRoot,
            "--effort", NormalizeEffort(effort),
            "--timeout-sec", Math.Clamp(timeoutSeconds, 120, 7200).ToString(System.Globalization.CultureInfo.InvariantCulture),
        ];

        if (limit > 0)
        {
            args.Add("--limit");
            args.Add(limit.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        return await RunPythonAsync(repoRoot, script, args, Math.Clamp(timeoutSeconds + 60, 180, 7300), ct);
    }

    public async Task<CurrentProblemProcessResult> RunAnalysisAsync(
        string dbPath,
        CurrentProblemRunRequest request,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(request.Problem))
            throw new ArgumentException("Problem is required.", nameof(request));

        Directory.CreateDirectory(SampleDir);
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string script = Path.Combine(repoRoot, "ai_current_problem_analyze.py");

        List<string> args =
        [
            "--sample-dir", SampleDir,
            "--db", dbPath,
            "--problem", request.Problem.Trim(),
            "--top", request.Top.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--candidate-limit", request.CandidateLimit.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--measure-report-limit", request.MeasureReportLimit.ToString(System.Globalization.CultureInfo.InvariantCulture),
            "--effort", NormalizeEffort(request.Effort),
            "--measure-effort", NormalizeEffort(request.MeasureEffort),
            "--html-effort", NormalizeEffort(request.HtmlEffort),
            "--timeout-sec", request.TimeoutSeconds.ToString(System.Globalization.CultureInfo.InvariantCulture),
        ];

        if (!string.IsNullOrWhiteSpace(request.Model))
        {
            args.Add("--model");
            args.Add(request.Model.Trim());
        }

        if (request.SkipMeasurements)
            args.Add("--skip-measurements");

        return await RunPythonAsync(repoRoot, script, args, request.TimeoutSeconds + 30, ct);
    }

    private async Task<CurrentProblemProcessResult> RunPythonAsync(
        string workDir,
        string script,
        IReadOnlyList<string> args,
        int timeoutSeconds,
        CancellationToken ct)
    {
        if (!File.Exists(script))
            throw new FileNotFoundException($"Python script not found: {script}", script);

        await _processLock.WaitAsync(ct);
        try
        {
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(Math.Clamp(timeoutSeconds, 30, 7200)));

            var psi = new ProcessStartInfo
            {
                FileName = "python",
                WorkingDirectory = workDir,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            psi.ArgumentList.Add(script);
            foreach (string arg in args)
                psi.ArgumentList.Add(arg);

            using var proc = new Process { StartInfo = psi, EnableRaisingEvents = true };
            var output = new StringBuilder();
            var error = new StringBuilder();
            proc.OutputDataReceived += (_, e) => { if (e.Data is not null) output.AppendLine(e.Data); };
            proc.ErrorDataReceived += (_, e) => { if (e.Data is not null) error.AppendLine(e.Data); };

            proc.Start();
            proc.BeginOutputReadLine();
            proc.BeginErrorReadLine();
            await proc.WaitForExitAsync(timeoutCts.Token);
            string outText = output.ToString();
            string errText = error.ToString();

            (string jsonPath, string htmlPath) = ExtractResultPaths(outText);
            return new CurrentProblemProcessResult(
                proc.ExitCode == 0,
                proc.ExitCode,
                outText,
                errText,
                jsonPath,
                htmlPath);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return new CurrentProblemProcessResult(false, -1, "", "Process timed out.", "", "");
        }
        finally
        {
            _processLock.Release();
        }
    }

    private (int Done, int Total, string UpdatedAt) ReadClassificationSummary()
    {
        string path = Path.Combine(DataRoot, "classification_summary.json");
        if (!File.Exists(path)) return (0, 0, "");
        try
        {
            using JsonDocument doc = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            JsonElement root = doc.RootElement;
            return (
                JsonInt(root, "done"),
                JsonInt(root, "total"),
                JsonString(root, "updatedAt"));
        }
        catch
        {
            return (0, 0, "");
        }
    }

    private static int CountJsonArray(string path)
    {
        if (!File.Exists(path)) return 0;
        try
        {
            using JsonDocument doc = JsonDocument.Parse(File.ReadAllText(path, Encoding.UTF8));
            return doc.RootElement.ValueKind == JsonValueKind.Array ? doc.RootElement.GetArrayLength() : 0;
        }
        catch
        {
            return 0;
        }
    }

    private static int CountMarkdownHeadings(string path)
    {
        if (!File.Exists(path)) return 0;
        return File.ReadLines(path, Encoding.UTF8).Count(line => line.StartsWith("### ", StringComparison.Ordinal));
    }

    private static string ResolveDataRoot(IWebHostEnvironment env)
    {
        string repoRoot = AiPromptRegistry.FindRepositoryRoot();
        string projectFile = Path.Combine(repoRoot, "JinoSupporter.Web", "JinoSupporter.Web.csproj");
        string projectDataRoot = Path.Combine(repoRoot, "JinoSupporter.Web", "App_Data", "ai-current-problem");

        if (File.Exists(projectFile) || Directory.Exists(projectDataRoot))
            return projectDataRoot;

        return Path.Combine(env.ContentRootPath, "App_Data", "ai-current-problem");
    }

    private static string LatestFile(string directory, string pattern)
    {
        if (!Directory.Exists(directory)) return "";
        return Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly)
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault() ?? "";
    }

    private static (string JsonPath, string HtmlPath) ExtractResultPaths(string output)
    {
        string text = (output ?? "").Trim();
        if (string.IsNullOrWhiteSpace(text)) return ("", "");

        int start = text.LastIndexOf('{');
        if (start < 0) return ("", "");

        try
        {
            using JsonDocument doc = JsonDocument.Parse(text[start..]);
            JsonElement root = doc.RootElement;
            return (JsonString(root, "json"), JsonString(root, "html"));
        }
        catch
        {
            return ("", "");
        }
    }

    private static int JsonInt(JsonElement root, string name)
        => root.TryGetProperty(name, out JsonElement value) && value.TryGetInt32(out int result) ? result : 0;

    private static double JsonDouble(JsonElement root, string name)
        => root.TryGetProperty(name, out JsonElement value) && value.TryGetDouble(out double result) ? result : 0;

    private static string JsonString(JsonElement root, string name)
        => root.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? ""
            : "";

    private static IReadOnlyList<string> JsonStringArray(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out JsonElement value))
            return [];

        if (value.ValueKind == JsonValueKind.String)
        {
            string single = value.GetString() ?? "";
            return string.IsNullOrWhiteSpace(single) ? [] : [single.Trim()];
        }

        if (value.ValueKind != JsonValueKind.Array)
            return [];

        return value.EnumerateArray()
            .Select(x => x.ValueKind == JsonValueKind.String ? x.GetString() ?? "" : x.ToString())
            .Select(x => x.Trim())
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static string NormalizeEffort(string? value)
    {
        string effort = (value ?? "").Trim().ToLowerInvariant();
        return effort is "minimal" or "low" or "medium" or "high" or "xhigh" ? effort : "medium";
    }

    private static void CopyDirectory(string source, string destination, CancellationToken ct)
    {
        Directory.CreateDirectory(destination);
        foreach (string dir in Directory.EnumerateDirectories(source, "*", SearchOption.AllDirectories))
        {
            ct.ThrowIfCancellationRequested();
            string relative = Path.GetRelativePath(source, dir);
            Directory.CreateDirectory(Path.Combine(destination, relative));
        }

        foreach (string file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
        {
            ct.ThrowIfCancellationRequested();
            string relative = Path.GetRelativePath(source, file);
            string target = Path.Combine(destination, relative);
            string? targetDir = Path.GetDirectoryName(target);
            if (!string.IsNullOrWhiteSpace(targetDir)) Directory.CreateDirectory(targetDir);
            File.Copy(file, target, overwrite: true);
        }
    }

    private static bool PathEquals(string a, string b)
        => string.Equals(
            Path.GetFullPath(a).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            Path.GetFullPath(b).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase);

    private static bool IsSubPathOf(string path, string possibleParent)
    {
        string fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string fullParent = Path.GetFullPath(possibleParent).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        return fullPath.StartsWith(fullParent, StringComparison.OrdinalIgnoreCase);
    }
}
