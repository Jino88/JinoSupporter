using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

public sealed record InputDataTestAutoPromptPayload(
    string Key,
    string Title,
    string Description,
    string SourceLabel,
    string Content);

public sealed record InputDataTestAnalysisParameters(
    string ReviewPurpose,
    IReadOnlyList<string> Tags,
    string Purpose,
    string PurposeCode,
    IReadOnlyList<string> TargetDefects,
    IReadOnlyList<string> ReviewItems,
    string Model,
    string Date,
    double? Confidence)
{
    public static InputDataTestAnalysisParameters Empty { get; } = new(
        "",
        Array.Empty<string>(),
        "",
        "",
        Array.Empty<string>(),
        Array.Empty<string>(),
        "",
        "",
        null);

    public bool HasAny =>
        !string.IsNullOrWhiteSpace(ReviewPurpose)
        || Tags.Count > 0
        || !string.IsNullOrWhiteSpace(Purpose)
        || !string.IsNullOrWhiteSpace(PurposeCode)
        || TargetDefects.Count > 0
        || ReviewItems.Count > 0
        || !string.IsNullOrWhiteSpace(Model)
        || !string.IsNullOrWhiteSpace(Date)
        || Confidence.HasValue;
}

public sealed record InputDataTestAnalysisResult(
    string AnalysisText,
    string AnalysisHtml,
    string Error,
    InputDataTestAnalysisParameters Parameters)
{
    public InputDataTestAnalysisResult(string analysisText, string analysisHtml, string error)
        : this(analysisText, analysisHtml, error, InputDataTestAnalysisParameters.Empty)
    {
    }
}

public static class InputDataTestAnalysisSupport
{
    public const string DetailedReviewIndexPath = @"D:\000. MyWorks\test\result\InputDataFinish\_review_html_detailed\review_index.html";

    public static string BuildCodexPrompt(string requestPath)
    {
        return AiPromptRegistry.Render(
            "input-data/analysis-runner.md",
            ("requestPath", requestPath),
            ("aiPreAnalysisPrompt", BuildAiPreAnalysisPrompt()),
            ("defaultAnalysisPrompt", BuildDefaultAnalysisPrompt()),
            ("visualizationSelectionPrompt", BuildVisualizationSelectionPrompt()),
            ("reviewIndexPrompt", BuildReviewIndexPrompt(requestPath)),
            ("workbookEvidencePrompt", BuildWorkbookEvidencePrompt()),
            ("analysisOutputPrompt", BuildAnalysisOutputPrompt()));
    }

    public static string BuildAiPreAnalysisPrompt()
        => AiPromptRegistry.Read("input-data/ai-pre-analysis.md");

    public static string BuildDefaultAnalysisPrompt()
        => AiPromptRegistry.Read("input-data/default-analysis.md");

    public static string BuildVisualizationSelectionPrompt()
        => AiPromptRegistry.Read("input-data/visualization-selection.md");

    public static string BuildReviewIndexPrompt(string requestPath)
        => AiPromptRegistry.Render(
            "input-data/review-index.md",
            ("requestPath", requestPath),
            ("detailedReviewIndexPath", DetailedReviewIndexPath));

    public static string BuildWorkbookEvidencePrompt()
        => AiPromptRegistry.Read("input-data/workbook-evidence.md");

    public static string BuildAnalysisOutputPrompt()
        => AiPromptRegistry.Read("input-data/analysis-output.md");

    public static JsonObject ParametersToJsonObject(InputDataTestAnalysisParameters? parameters)
    {
        InputDataTestAnalysisParameters p = parameters ?? InputDataTestAnalysisParameters.Empty;
        return new JsonObject
        {
            ["reviewPurpose"] = p.ReviewPurpose,
            ["tags"] = JsonSerializer.SerializeToNode(p.Tags.ToArray()),
            ["purpose"] = p.Purpose,
            ["purposeCode"] = p.PurposeCode,
            ["targetDefects"] = JsonSerializer.SerializeToNode(p.TargetDefects.ToArray()),
            ["reviewItems"] = JsonSerializer.SerializeToNode(p.ReviewItems.ToArray()),
            ["model"] = p.Model,
            ["date"] = p.Date,
            ["confidence"] = p.Confidence.HasValue ? JsonValue.Create(p.Confidence.Value) : null
        };
    }

    public static IReadOnlyList<string> BuildReviewIndexMatchKeys(params string?[] values)
    {
        var keys = new List<string>();
        foreach (string? value in values)
        {
            AddReviewIndexMatchKey(keys, value);

            string fileName = FileNameFromMaybePath(value);
            AddReviewIndexMatchKey(keys, fileName);

            if (!string.IsNullOrWhiteSpace(fileName))
                AddReviewIndexMatchKey(keys, Path.GetFileNameWithoutExtension(fileName));

            AddReviewIndexMatchKey(keys, NormalizeReviewIndexMatchKey(value));
            AddReviewIndexMatchKey(keys, NormalizeReviewIndexMatchKey(fileName));
        }

        return keys.Take(24).ToArray();
    }

    private static void AddReviewIndexMatchKey(List<string> keys, string? value)
    {
        string key = Regex.Replace(value ?? "", @"\s+", " ").Trim();
        if (key.Length < 2) return;
        if (key.Length > 220) key = key[..220].Trim();
        if (!keys.Contains(key, StringComparer.OrdinalIgnoreCase))
            keys.Add(key);
    }

    private static string FileNameFromMaybePath(string? value)
    {
        string text = (value ?? "").Trim();
        if (text.Length == 0) return "";

        try
        {
            string normalized = text.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            string fileName = Path.GetFileName(normalized);
            return string.IsNullOrWhiteSpace(fileName) ? text : fileName;
        }
        catch
        {
            return text;
        }
    }

    private static string NormalizeReviewIndexMatchKey(string? value)
    {
        string text = FileNameFromMaybePath(value);
        if (text.Length == 0) return "";

        text = Regex.Replace(text, @"(?i)\.(xlsx|xlsm|xlsb|xls)$", "");
        text = Regex.Replace(text, @"(?i)(?:_clean_textonly|_textonly|_clean)$", "");
        for (int i = 0; i < 4; i++)
            text = Regex.Replace(text, @"(?i)(?:_[0-9a-f]{8,}|_\d{8,})$", "");

        text = Regex.Replace(text, @"\s+", " ").Trim(' ', '.', '_', '-');
        return text;
    }

    public static IReadOnlyList<InputDataTestAutoPromptPayload> LoadEnabledAutoPrompts(string repoRoot)
    {
        string libraryPath = AutoPromptLibraryPath();
        if (!File.Exists(libraryPath))
        {
            string legacy = Path.Combine(repoRoot, "tmp", "input-data-test", "auto-prompts.json");
            if (File.Exists(legacy)) libraryPath = legacy;
        }

        if (!File.Exists(libraryPath)) return [];

        try
        {
            string json = File.ReadAllText(libraryPath);
            var records = JsonSerializer.Deserialize<List<AutoPromptStorage>>(
                json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];

            return records
                .Where(r => r.Enabled && !string.IsNullOrWhiteSpace(r.Content))
                .Select(r => new InputDataTestAutoPromptPayload(
                    r.Key ?? "",
                    string.IsNullOrWhiteSpace(r.Title) ? r.Key ?? "" : r.Title,
                    r.Description ?? "",
                    string.IsNullOrWhiteSpace(r.SourceLabel) ? "Auto Prompt" : r.SourceLabel,
                    r.Content ?? ""))
                .ToArray();
        }
        catch
        {
            return [];
        }
    }

    public static InputDataTestAnalysisResult ParseStepResult(string raw)
    {
        string text = ExtractStepResultJson(raw);
        if (!string.IsNullOrWhiteSpace(text) && TryParseStepResultJson(text, out InputDataTestAnalysisResult parsed))
        {
            return parsed;
        }

        string fallbackText = string.IsNullOrWhiteSpace(raw)
            ? "(Codex returned an empty result.)"
            : raw.Trim();
        return new InputDataTestAnalysisResult(
            fallbackText,
            "",
            "Codex output was not valid JSON with analysisHtml. Text was kept, but no AI-generated HTML was displayed.");
    }

    public static string ExtractStepResultJson(string raw)
    {
        string text = CleanJsonText(raw);
        if (string.IsNullOrWhiteSpace(text)) return "";

        if (TryParseStepResultJson(text, out _))
            return text;

        string fallback = "";
        foreach (string candidate in EnumerateJsonObjectCandidates(text))
        {
            if (!TryParseStepResultJson(candidate, out InputDataTestAnalysisResult result))
                continue;

            if (!string.IsNullOrWhiteSpace(result.AnalysisHtml))
                fallback = candidate;
            else if (fallback.Length == 0)
                fallback = candidate;
        }

        return fallback;
    }

    private static string AutoPromptLibraryPath()
    {
        string dir = AppStoragePaths.Combine("InputDataTest");
        return Path.Combine(dir, "auto-prompts.json");
    }

    private static string CleanJsonText(string raw)
    {
        string text = (raw ?? "").Trim();
        if (!text.StartsWith("```", StringComparison.Ordinal)) return text;

        int firstNewline = text.IndexOf('\n');
        if (firstNewline >= 0) text = text[(firstNewline + 1)..];
        int fence = text.LastIndexOf("```", StringComparison.Ordinal);
        if (fence >= 0) text = text[..fence];
        return text.Trim();
    }

    private static bool TryParseStepResultJson(string text, out InputDataTestAnalysisResult result)
    {
        result = new InputDataTestAnalysisResult("", "", "");

        try
        {
            using JsonDocument doc = JsonDocument.Parse(text);
            JsonElement root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return false;

            string analysisText = JsonString(root, "analysisText", "text", "result", "analysisMarkdown");
            string analysisHtml = JsonString(root, "analysisHtml", "html", "resultHtml").Trim();
            if (string.IsNullOrWhiteSpace(analysisText) && string.IsNullOrWhiteSpace(analysisHtml))
                return false;

            InputDataTestAnalysisParameters parameters = ParseParameters(root);
            string error = string.IsNullOrWhiteSpace(analysisHtml)
                ? "Codex returned analysisText but did not return analysisHtml. Text was kept, but no AI-generated HTML was displayed."
                : "";

            result = new InputDataTestAnalysisResult(
                analysisText.Trim(),
                analysisHtml,
                error,
                parameters);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static InputDataTestAnalysisParameters ParseParameters(JsonElement root)
    {
        JsonElement source = root;
        if (TryGetPropertyIgnoreCase(root, "parameters", out JsonElement parameters)
            || TryGetPropertyIgnoreCase(root, "frontParameters", out parameters)
            || TryGetPropertyIgnoreCase(root, "parameter", out parameters))
        {
            if (parameters.ValueKind == JsonValueKind.Object)
                source = parameters;
        }

        if (source.ValueKind != JsonValueKind.Object)
            return InputDataTestAnalysisParameters.Empty;

        var parsed = new InputDataTestAnalysisParameters(
            CleanParameterText(JsonString(source, "reviewPurpose", "reviewObjective", "objective")),
            JsonStringList(source, "tags", "reviewTags", "checkTags"),
            CleanParameterText(JsonString(source, "purpose", "testPurpose", "workbookPurpose")),
            CleanParameterText(JsonString(source, "purposeCode", "purposeCategory", "purposeNo")),
            JsonStringList(source, "targetDefects", "targetDefect", "defects", "targetDefectNames"),
            JsonStringList(source, "reviewItems", "reviewItem", "checkItems", "processes"),
            CleanParameterText(JsonString(source, "model", "productModel")),
            CleanParameterText(JsonString(source, "date", "reportDate", "testDate")),
            JsonDouble(source, "confidence", "parameterConfidence"));

        return parsed.HasAny ? parsed : InputDataTestAnalysisParameters.Empty;
    }

    private static IEnumerable<string> EnumerateJsonObjectCandidates(string text)
    {
        for (int start = 0; start < text.Length; start++)
        {
            if (text[start] != '{') continue;

            int probe = start + 1;
            while (probe < text.Length && char.IsWhiteSpace(text[probe])) probe++;
            if (probe >= text.Length || text[probe] != '"') continue;

            bool inString = false;
            bool escaped = false;
            int depth = 0;

            for (int i = start; i < text.Length; i++)
            {
                char c = text[i];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                    }
                    else if (c == '\\')
                    {
                        escaped = true;
                    }
                    else if (c == '"')
                    {
                        inString = false;
                    }
                    continue;
                }

                if (c == '"')
                {
                    inString = true;
                }
                else if (c == '{')
                {
                    depth++;
                }
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        yield return text[start..(i + 1)];
                        break;
                    }
                }
            }
        }
    }

    private static string JsonString(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetPropertyIgnoreCase(root, name, out JsonElement value)) continue;
            string text = value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString();
            if (!string.IsNullOrWhiteSpace(text)) return text;
        }
        return "";
    }

    private static IReadOnlyList<string> JsonStringList(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetPropertyIgnoreCase(root, name, out JsonElement value)) continue;

            IEnumerable<string> values = value.ValueKind switch
            {
                JsonValueKind.Array => value.EnumerateArray().Select(JsonElementToString),
                JsonValueKind.String => Regex.Split(value.GetString() ?? "", @"[,;\n|]+"),
                _ => [value.ToString()]
            };

            string[] result = values
                .Select(CleanParameterText)
                .Where(v => v.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(12)
                .ToArray();
            if (result.Length > 0) return result;
        }

        return Array.Empty<string>();
    }

    private static string JsonElementToString(JsonElement value)
        => value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString();

    private static double? JsonDouble(JsonElement root, params string[] names)
    {
        foreach (string name in names)
        {
            if (!TryGetPropertyIgnoreCase(root, name, out JsonElement value)) continue;
            double? parsed = value.ValueKind switch
            {
                JsonValueKind.Number when value.TryGetDouble(out double numericValue) => numericValue,
                JsonValueKind.String when double.TryParse(value.GetString(), out double stringValue) => stringValue,
                _ => null
            };
            if (!parsed.HasValue) continue;

            double confidence = parsed.Value;
            if (confidence > 1 && confidence <= 100) confidence /= 100;
            if (confidence is >= 0 and <= 1) return confidence;
        }
        return null;
    }

    private static bool TryGetPropertyIgnoreCase(JsonElement root, string name, out JsonElement value)
    {
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty(name, out value))
            return true;

        if (root.ValueKind == JsonValueKind.Object)
        {
            foreach (JsonProperty property in root.EnumerateObject())
            {
                if (property.NameEquals(name)
                    || string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string CleanParameterText(string value)
    {
        string text = Regex.Replace(value ?? "", @"\s+", " ").Trim(' ', '-', '*', ':');
        if (text.Equals("unknown", StringComparison.OrdinalIgnoreCase)
            || text.Equals("n/a", StringComparison.OrdinalIgnoreCase)
            || text.Equals("na", StringComparison.OrdinalIgnoreCase)
            || text.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }
        return text.Length > 240 ? text[..240].Trim() : text;
    }

    private sealed record AutoPromptStorage(
        string? Key,
        string? Title,
        string? Description,
        string? SourceLabel,
        string? Content,
        bool Enabled,
        bool IsCustom);
}
