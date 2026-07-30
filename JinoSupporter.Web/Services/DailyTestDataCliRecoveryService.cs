using System.Text.Json;

namespace JinoSupporter.Web.Services;

public sealed class DailyTestDataCliRecoveryService(
    IWebHostEnvironment env,
    WebRepository repo,
    AppActivityLogger activity) : BackgroundService
{
    private static readonly TimeSpan ScanInterval = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan CompletedFileGrace = TimeSpan.FromSeconds(45);

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        await Task.Delay(TimeSpan.FromSeconds(15), stoppingToken);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await RecoverCompletedRunsAsync(stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                return;
            }
            catch (Exception ex)
            {
                activity.Log("DailyTestRecovery", $"scan failed: {ex.Message}");
            }

            await Task.Delay(ScanInterval, stoppingToken);
        }
    }

    private async Task RecoverCompletedRunsAsync(CancellationToken token)
    {
        string tmpDir = Path.Combine(FindRepoRoot(), "tmp");
        if (!Directory.Exists(tmpDir)) return;

        foreach (string donePath in Directory.EnumerateFiles(tmpDir, "daily_test_analysis_done_full_*.json"))
        {
            token.ThrowIfCancellationRequested();
            await TryRecoverFullRunAsync(tmpDir, donePath, token);
        }

        foreach (string donePath in Directory.EnumerateFiles(tmpDir, "daily_test_analysis_done_added_*.json"))
        {
            token.ThrowIfCancellationRequested();
            await TryRecoverAddedRunAsync(tmpDir, donePath, token);
        }
    }

    private async Task TryRecoverFullRunAsync(string tmpDir, string donePath, CancellationToken token)
    {
        if (!IsOldEnough(donePath)) return;

        string runId = RunIdFromDonePath(donePath, "daily_test_analysis_done_full_");
        if (string.IsNullOrWhiteSpace(runId)) return;

        string outputPath = Path.Combine(tmpDir, $"daily_test_analysis_output_full_{runId}.txt");
        string requestPath = Path.Combine(tmpDir, $"daily_test_analysis_request_full_{runId}.json");
        if (!File.Exists(outputPath) || !File.Exists(requestPath)) return;

        DailyTestRequest? request = ReadRequest(requestPath);
        if (request is null || request.ItemId <= 0) return;

        DailyTestAiResult result = ParseAiResult(await File.ReadAllTextAsync(outputPath, token));
        if (!result.HasRequiredResult) return;

        string itemDataText = string.IsNullOrWhiteSpace(request.CumulativeDataText)
            ? request.DataText
            : request.CumulativeDataText;
        string itemPromptText = string.IsNullOrWhiteSpace(request.CumulativePromptText)
            ? request.PromptText
            : request.CumulativePromptText;

        repo.SaveDailyTestDataAnalysisResult(
            request.ItemId,
            itemDataText,
            itemPromptText,
            result.ParametersJson,
            result.AnalysisMarkdown,
            result.AnalysisHtml);
        string writtenHtmlPath = TryWriteRequestedStandaloneHtmlFile(
            request.StandaloneOutputPath,
            itemPromptText,
            result.AnalysisHtml);

        DeleteRunFiles(requestPath, outputPath, donePath);
        activity.Log("DailyTestRecovery", $"saved missed full result item={request.ItemId} run={runId} standaloneHtml={writtenHtmlPath}");
    }

    private async Task TryRecoverAddedRunAsync(string tmpDir, string donePath, CancellationToken token)
    {
        if (!IsOldEnough(donePath)) return;

        string runId = RunIdFromDonePath(donePath, "daily_test_analysis_done_added_");
        if (string.IsNullOrWhiteSpace(runId)) return;

        string outputPath = Path.Combine(tmpDir, $"daily_test_analysis_output_added_{runId}.txt");
        string requestPath = Path.Combine(tmpDir, $"daily_test_analysis_request_added_{runId}.json");
        if (!File.Exists(outputPath) || !File.Exists(requestPath)) return;

        DailyTestRequest? request = ReadRequest(requestPath);
        if (request is null || request.ItemId <= 0) return;

        DailyTestAiResult result = ParseAiResult(await File.ReadAllTextAsync(outputPath, token));
        if (!result.HasRequiredResult) return;

        string historyParameters = string.IsNullOrWhiteSpace(result.HistoryParametersJson)
            ? result.ParametersJson
            : result.HistoryParametersJson;
        string historyMarkdown = string.IsNullOrWhiteSpace(result.HistoryAnalysisMarkdown)
            ? result.AnalysisMarkdown
            : result.HistoryAnalysisMarkdown;
        string historyHtml = string.IsNullOrWhiteSpace(result.HistoryAnalysisHtml)
            ? result.AnalysisHtml
            : result.HistoryAnalysisHtml;

        bool insertedHistory = EnsureHistorySaved(
            request.ItemId,
            request.Name,
            request.DataText,
            request.PromptText,
            historyParameters,
            historyMarkdown,
            historyHtml);

        DailyTestDataItemRecord? item = repo.GetDailyTestDataItem(request.ItemId);
        bool firstSegment = request.HistorySnapshotCount == 0
                            && string.IsNullOrWhiteSpace(request.CumulativeDataText)
                            && string.IsNullOrWhiteSpace(item?.AnalysisMarkdown)
                            && string.IsNullOrWhiteSpace(item?.AnalysisHtml);
        if (firstSegment)
        {
            repo.SaveDailyTestDataAnalysisResult(
                request.ItemId,
                request.DataText,
                request.PromptText,
                result.ParametersJson,
                result.AnalysisMarkdown,
                result.AnalysisHtml);
        }

        DeleteRunFiles(requestPath, outputPath, donePath);
        activity.Log("DailyTestRecovery", $"saved missed added result item={request.ItemId} run={runId} historyInserted={insertedHistory} firstSegment={firstSegment}");
    }

    private bool EnsureHistorySaved(
        long itemId,
        string name,
        string dataText,
        string promptText,
        string parametersJson,
        string analysisMarkdown,
        string analysisHtml)
    {
        bool exists = repo.GetDailyTestDataHistory(itemId).Any(h =>
            string.Equals(SnapshotKey(h.DataText), SnapshotKey(dataText), StringComparison.Ordinal)
            && string.Equals(SnapshotKey(h.PromptText), SnapshotKey(promptText), StringComparison.Ordinal));
        if (exists) return false;

        repo.SaveDailyTestDataHistory(
            itemId,
            name,
            dataText,
            promptText,
            parametersJson,
            analysisMarkdown,
            analysisHtml,
            DateTime.UtcNow.ToString("o"));
        return true;
    }

    private static DailyTestRequest? ReadRequest(string requestPath)
    {
        try
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(requestPath));
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) return null;

            return new DailyTestRequest(
                JsonLong(root, "itemId"),
                JsonString(root, "name"),
                JsonString(root, "dataText"),
                JsonString(root, "promptText"),
                JsonString(root, "workflowPhase"),
                JsonString(root, "cumulativeDataText"),
                JsonString(root, "cumulativePromptText"),
                JsonString(root, "standaloneOutputPath"),
                JsonInt(root, "historySnapshotCount"));
        }
        catch
        {
            return null;
        }
    }

    private static DailyTestAiResult ParseAiResult(string raw)
    {
        string json = ExtractJson(raw);
        try
        {
            using var doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;
            string parameters = RawJsonProp(root, "parameters");
            if (string.IsNullOrWhiteSpace(parameters)) parameters = "{}";
            string analysis = JsonString(root, "analysisMarkdown");
            string html = NormalizeAnalysisHtmlForStorage(JsonString(root, "analysisHtml"));
            if (string.IsNullOrWhiteSpace(html) && !string.IsNullOrWhiteSpace(analysis))
                html = ConvertAnalysisMarkdownToHtml(analysis);
            string historyAnalysis = JsonString(root,
                "historyAnalysisMarkdown",
                "currentAnalysisMarkdown",
                "currentInputAnalysisMarkdown");
            string historyHtml = SanitizeHtmlFragment(JsonString(root,
                "historyAnalysisHtml",
                "currentAnalysisHtml",
                "currentInputAnalysisHtml"));
            if (string.IsNullOrWhiteSpace(historyHtml) && !string.IsNullOrWhiteSpace(historyAnalysis))
                historyHtml = ConvertAnalysisMarkdownToHtml(historyAnalysis);

            return new DailyTestAiResult(
                parameters,
                analysis,
                html,
                RawJsonProp(root, "historyParameters", "currentParameters", "currentInputParameters"),
                historyAnalysis,
                historyHtml);
        }
        catch
        {
            return new DailyTestAiResult("", "", "", "", "", "");
        }
    }

    private static string ExtractJson(string raw)
    {
        raw = (raw ?? "").Trim();
        if (raw.StartsWith("```", StringComparison.Ordinal))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```", StringComparison.Ordinal))
                raw = raw[..raw.LastIndexOf("```", StringComparison.Ordinal)];
        }

        const string resultStart = "{\"parameters\"";
        for (int searchEnd = raw.Length - 1; searchEnd >= 0;)
        {
            int resultOpen = raw.LastIndexOf(resultStart, searchEnd, StringComparison.Ordinal);
            if (resultOpen < 0) break;
            if (TryExtractBalancedJsonObject(raw, resultOpen, out string candidate) && IsAiResultJson(candidate))
                return candidate;
            if (resultOpen == 0) break;
            searchEnd = resultOpen - 1;
        }

        for (int open = raw.LastIndexOf('{'); open >= 0;)
        {
            if (TryExtractBalancedJsonObject(raw, open, out string candidate) && IsAiResultJson(candidate))
                return candidate;
            if (open == 0) break;
            open = raw.LastIndexOf('{', open - 1);
        }

        int firstOpen = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        return firstOpen >= 0 && close > firstOpen ? raw[firstOpen..(close + 1)] : raw;
    }

    private static bool TryExtractBalancedJsonObject(string text, int open, out string json)
    {
        json = "";
        if (open < 0 || open >= text.Length || text[open] != '{') return false;

        bool inString = false;
        bool escape = false;
        int depth = 0;
        for (int i = open; i < text.Length; i++)
        {
            char ch = text[i];
            if (inString)
            {
                if (escape)
                {
                    escape = false;
                }
                else if (ch == '\\')
                {
                    escape = true;
                }
                else if (ch == '"')
                {
                    inString = false;
                }
                continue;
            }

            if (ch == '"')
            {
                inString = true;
                continue;
            }

            if (ch == '{')
            {
                depth++;
            }
            else if (ch == '}')
            {
                depth--;
                if (depth == 0)
                {
                    json = text[open..(i + 1)];
                    return true;
                }
                if (depth < 0) return false;
            }
        }

        return false;
    }

    private static bool IsAiResultJson(string value)
    {
        try
        {
            using var doc = JsonDocument.Parse(value);
            JsonElement root = doc.RootElement;
            return root.ValueKind == JsonValueKind.Object
                   && root.TryGetProperty("parameters", out _)
                   && root.TryGetProperty("analysisMarkdown", out _)
                   && root.TryGetProperty("analysisHtml", out _);
        }
        catch
        {
            return false;
        }
    }

    private static string RawJsonProp(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!root.TryGetProperty(name, out JsonElement value)) continue;
            if (value.ValueKind is JsonValueKind.Object or JsonValueKind.Array) return value.GetRawText();
            string text = value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString();
            if (!string.IsNullOrWhiteSpace(text)) return text;
        }
        return "";
    }

    private static string JsonString(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!root.TryGetProperty(name, out JsonElement value)) continue;
            string text = value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString();
            if (!string.IsNullOrWhiteSpace(text)) return text;
        }
        return "";
    }

    private static long JsonLong(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out JsonElement value)) return 0;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out long n)) return n;
        return long.TryParse(value.ToString(), out n) ? n : 0;
    }

    private static int JsonInt(JsonElement root, string name)
    {
        if (!root.TryGetProperty(name, out JsonElement value)) return 0;
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out int n)) return n;
        return int.TryParse(value.ToString(), out n) ? n : 0;
    }

    private static string ConvertAnalysisMarkdownToHtml(string? markdown)
    {
        string text = (markdown ?? "").Trim();
        if (string.IsNullOrWhiteSpace(text)) return "";
        string encoded = System.Net.WebUtility.HtmlEncode(text)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n')
            .Replace("\n", "<br>", StringComparison.Ordinal);
        return $"<p>{encoded}</p>";
    }

    private static string SanitizeHtmlFragment(string html)
    {
        string text = (html ?? "").Trim();
        if (string.IsNullOrWhiteSpace(text)) return "";

        var body = System.Text.RegularExpressions.Regex.Match(
            text,
            @"<body\b[^>]*>(.*?)</body>",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
        if (body.Success) text = body.Groups[1].Value;

        text = System.Text.RegularExpressions.Regex.Replace(
            text,
            @"<\s*(script|iframe|object|embed|style|link|meta|base)\b[^>]*>.*?<\s*/\s*\1\s*>",
            "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
        text = System.Text.RegularExpressions.Regex.Replace(
            text,
            @"<\s*(script|iframe|object|embed|style|link|meta|base)\b[^>]*/?\s*>",
            "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(
            text,
            @"\s+on\w+\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)",
            "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        text = System.Text.RegularExpressions.Regex.Replace(
            text,
            @"\s+(href|src)\s*=\s*(['""])\s*javascript:[^'""]*\2",
            " $1=\"#\"",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        return text;
    }

    private static string NormalizeAnalysisHtmlForStorage(string html)
    {
        string text = StripMarkdownHtmlFence(html ?? "").Trim();
        if (string.IsNullOrWhiteSpace(text)) return "";
        return LooksLikeStandaloneHtmlDocument(text) ? text : SanitizeHtmlFragment(text);
    }

    private static bool LooksLikeStandaloneHtmlDocument(string html)
    {
        string text = (html ?? "").TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
        if (text.StartsWith("```", StringComparison.Ordinal))
            text = StripMarkdownHtmlFence(text).TrimStart('\uFEFF', ' ', '\t', '\r', '\n');
        return text.StartsWith("<!doctype", StringComparison.OrdinalIgnoreCase)
               || System.Text.RegularExpressions.Regex.IsMatch(
                   text,
                   @"<html\b",
                   System.Text.RegularExpressions.RegexOptions.IgnoreCase);
    }

    private static string StripMarkdownHtmlFence(string raw)
    {
        string text = (raw ?? "").Trim();
        if (!text.StartsWith("```", StringComparison.Ordinal)) return text;

        int nl = text.IndexOf('\n');
        if (nl >= 0) text = text[(nl + 1)..];
        if (text.TrimEnd().EndsWith("```", StringComparison.Ordinal))
            text = text[..text.LastIndexOf("```", StringComparison.Ordinal)];
        return text.Trim();
    }

    private static string TryWriteRequestedStandaloneHtmlFile(string standaloneOutputPath, string promptText, string analysisHtml)
    {
        if (!LooksLikeStandaloneHtmlDocument(analysisHtml)) return "";

        string outputPath = string.IsNullOrWhiteSpace(standaloneOutputPath)
            ? ExtractRequestedStandaloneOutputPath(promptText)
            : standaloneOutputPath;
        if (string.IsNullOrWhiteSpace(outputPath)) return "";

        try
        {
            string fullPath = Path.GetFullPath(outputPath);
            if (!fullPath.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
                && !fullPath.EndsWith(".htm", StringComparison.OrdinalIgnoreCase))
                return "";

            string? dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(dir)) Directory.CreateDirectory(dir);
            File.WriteAllText(
                fullPath,
                analysisHtml.Trim(),
                new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return fullPath;
        }
        catch
        {
            return "";
        }
    }

    private static string ExtractRequestedStandaloneOutputPath(string text)
    {
        foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(
                     text ?? "",
                     @"(?<path>(?:[A-Za-z]:\\|\\\\)[^\r\n""<>|?*]+?\.html?)",
                     System.Text.RegularExpressions.RegexOptions.IgnoreCase))
        {
            string path = match.Groups["path"].Value
                .Trim()
                .Trim('`', '\'', '"', ' ', '\t', '.', ',', ';', ')', ']', '}');
            if (path.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
                || path.EndsWith(".htm", StringComparison.OrdinalIgnoreCase))
                return path;
        }
        return "";
    }

    private static bool IsOldEnough(string path)
        => DateTime.UtcNow - File.GetLastWriteTimeUtc(path) >= CompletedFileGrace;

    private static string RunIdFromDonePath(string donePath, string prefix)
    {
        string name = Path.GetFileNameWithoutExtension(donePath);
        return name.StartsWith(prefix, StringComparison.Ordinal) ? name[prefix.Length..] : "";
    }

    private static string SnapshotKey(string text)
        => (text ?? "").Replace("\r\n", "\n").Replace('\r', '\n').Trim();

    private static void DeleteRunFiles(params string[] paths)
    {
        foreach (string path in paths)
        {
            try
            {
                if (File.Exists(path)) File.Delete(path);
            }
            catch { }
        }
    }

    private string FindRepoRoot()
    {
        string dir = env.ContentRootPath;
        for (int i = 0; i < 8; i++)
        {
            if (File.Exists(Path.Combine(dir, "JinoSupporter.sln"))) return dir;
            string? parent = Path.GetDirectoryName(dir.TrimEnd('\\', '/'));
            if (string.IsNullOrEmpty(parent) || parent == dir) break;
            dir = parent;
        }
        return env.ContentRootPath;
    }

    private sealed record DailyTestRequest(
        long ItemId,
        string Name,
        string DataText,
        string PromptText,
        string WorkflowPhase,
        string CumulativeDataText,
        string CumulativePromptText,
        string StandaloneOutputPath,
        int HistorySnapshotCount);

    private sealed record DailyTestAiResult(
        string ParametersJson,
        string AnalysisMarkdown,
        string AnalysisHtml,
        string HistoryParametersJson,
        string HistoryAnalysisMarkdown,
        string HistoryAnalysisHtml)
    {
        public bool HasRequiredResult =>
            !string.IsNullOrWhiteSpace(AnalysisHtml);
    }
}
