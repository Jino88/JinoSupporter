using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services;

public sealed record CodexTokenUsage(int? InputTokens, int? OutputTokens, int? TotalTokens);

public sealed record CodexTranslationResult(string SourceText, string Translation, CodexTokenUsage? Usage);

public sealed record CodexTextResult(string Text, string Model, CodexTokenUsage? Usage);

public sealed class CodexApiService(
    HttpClient http,
    IConfiguration config,
    WebRepository repo,
    AiProviderSettingsService providers)
{
    /// <summary>"Codex Luna" — the same model WorkHub runs for its light structure/summary work.</summary>
    public const string LunaModel = "gpt-5.6-luna";
    public const string DefaultTranslateModel = LunaModel;
    public const string DefaultReasoningEffort = "low";
    public static readonly string[] ReasoningEfforts = ["none", "low", "medium", "high"];

    public bool IsConfigured => providers.CodexApiEnabled && !string.IsNullOrWhiteSpace(ApiKey);

    public string Model => FirstNonBlank(
        repo.GetSetting("OpenAI:TranslateModel"),
        repo.GetSetting("Codex:TranslateModel"),
        Environment.GetEnvironmentVariable("OPENAI_TRANSLATE_MODEL"),
        config["OpenAI:TranslateModel"],
        config["Codex:TranslateModel"],
        DefaultTranslateModel);

    /// <summary>Reasoning effort sent with translation requests to reasoning models.</summary>
    public string TranslateEffort => FirstNonBlank(
        repo.GetSetting("OpenAI:TranslateEffort"),
        config["OpenAI:TranslateEffort"],
        DefaultReasoningEffort);

    private string ApiKey => FirstNonBlank(
        repo.GetSetting("OpenAI:ApiKey"),
        repo.GetSetting("Codex:ApiKey"),
        Environment.GetEnvironmentVariable("OPENAI_API_KEY"),
        config["OpenAI:ApiKey"],
        config["Codex:ApiKey"]);

    public async Task<CodexTranslationResult> TranslateAsync(
        string text,
        string targetLanguage,
        IReadOnlyList<(string MediaType, string Base64)> images,
        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Codex API is not configured.");

        var content = new JsonArray
        {
            new JsonObject
            {
                ["type"] = "input_text",
                ["text"] = BuildPrompt(text, targetLanguage, images.Count)
            }
        };

        foreach (var image in images)
        {
            if (string.IsNullOrWhiteSpace(image.Base64)) continue;
            string mediaType = string.IsNullOrWhiteSpace(image.MediaType) ? "image/png" : image.MediaType;
            content.Add(new JsonObject
            {
                ["type"] = "input_image",
                ["image_url"] = $"data:{mediaType};base64,{image.Base64}",
                ["detail"] = "high"
            });
        }

        var body = new JsonObject
        {
            ["model"] = Model,
            ["input"] = new JsonArray
            {
                new JsonObject
                {
                    ["role"] = "user",
                    ["content"] = content
                }
            },
            ["text"] = new JsonObject
            {
                ["format"] = BuildTranslationJsonSchema()
            },
            // Reasoning tokens count against this cap, so leave room beyond the translation itself.
            ["max_output_tokens"] = 8192,
            ["store"] = false
        };
        AddReasoning(body, Model, TranslateEffort);

        string raw = await PostResponsesAsync(body, ct);
        string output = ExtractOutputText(raw);
        if (string.IsNullOrWhiteSpace(output))
            throw new InvalidOperationException("Codex API returned an empty response.");

        CodexTranslationResult result = ParseTranslationResult(output);
        return result with { Usage = ExtractUsage(raw) };
    }

    /// <summary>
    /// One free-form Responses API call: <paramref name="content"/> holds the user message parts
    /// (input_text / input_image / input_file). Returns the output text and token usage.
    /// </summary>
    public async Task<CodexTextResult> GenerateAsync(
        string instructions,
        JsonArray content,
        string model,
        string effort,
        int maxOutputTokens,
        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Codex API is not configured. Set the API key in Setting.");

        var body = new JsonObject
        {
            ["model"] = model,
            ["instructions"] = instructions,
            ["input"] = new JsonArray
            {
                new JsonObject { ["role"] = "user", ["content"] = content }
            },
            ["max_output_tokens"] = maxOutputTokens,
            ["store"] = false
        };
        AddReasoning(body, model, effort);

        string raw = await PostResponsesAsync(body, ct);
        string output = ExtractOutputText(raw).Trim();
        if (output.Length == 0)
            throw new InvalidOperationException("Codex API returned an empty response.");
        return new CodexTextResult(output, model, ExtractUsage(raw));
    }

    /// <summary>Adds the call to the running Codex API token totals shown on the AI Usages page.</summary>
    public void RecordUsage(CodexTokenUsage? usage)
    {
        if (usage is null) return;
        long Read(string key) => long.TryParse(repo.GetSetting(key), out long v) ? v : 0;
        var inv = System.Globalization.CultureInfo.InvariantCulture;
        long input = Read("AiUsages:CodexApiInputTokens") + (usage.InputTokens ?? 0);
        long output = Read("AiUsages:CodexApiOutputTokens") + (usage.OutputTokens ?? 0);
        long total = Read("AiUsages:CodexApiTotalTokens") + (usage.TotalTokens ?? ((usage.InputTokens ?? 0) + (usage.OutputTokens ?? 0)));
        repo.SetSetting("AiUsages:CodexApiInputTokens", input.ToString(inv));
        repo.SetSetting("AiUsages:CodexApiOutputTokens", output.ToString(inv));
        repo.SetSetting("AiUsages:CodexApiTotalTokens", total.ToString(inv));
        repo.SetSetting("AiUsages:CodexApiLastRequestAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", inv));
    }

    private async Task<string> PostResponsesAsync(JsonObject body, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, "responses")
        {
            Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ApiKey);

        using HttpResponseMessage response = await http.SendAsync(request, ct);
        string raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"Codex API error {(int)response.StatusCode}: {raw}");
        return raw;
    }

    /// <summary>gpt-5.x and o-series take a reasoning effort; older chat models reject the field.</summary>
    private static void AddReasoning(JsonObject body, string model, string effort)
    {
        bool reasoningModel = model.StartsWith("gpt-5", StringComparison.OrdinalIgnoreCase)
                              || (model.Length > 1 && model[0] is 'o' or 'O' && char.IsDigit(model[1]));
        if (!reasoningModel || string.IsNullOrWhiteSpace(effort)) return;
        body["reasoning"] = new JsonObject { ["effort"] = effort.Trim().ToLowerInvariant() };
    }

    private const string TranslationPromptPath = "translation/openai-ocr-translation.md";

    // Same text as AI_PROMPTS/translation/openai-ocr-translation.md, used when that folder is not
    // deployed next to the app so translation does not fail on a missing template.
    private const string FallbackTranslationPrompt = """
        You are a professional AI OCR and translation engine.
        Target language: {{targetLanguage}}

        Instructions:
        - If image files are provided, first perform AI OCR yourself and transcribe the visible source text exactly, line by line, into sourceText.
        - Do not guess hidden or unreadable words from context. If a word or character is unreadable, write [unclear] in sourceText rather than inventing a term.
        - Do not substitute similar-looking words, brands, places, or technical terms. Preserve names, Korean syllables, abbreviations, and domain terms exactly in sourceText.
        - After sourceText is finalized, translate only that sourceText into the target language.
        - If source text is provided without images, use it verbatim as sourceText.
        - Preserve product codes, model names, serial numbers, measurements, paths, and part numbers exactly unless they are natural-language text.
        - If source text and images are both provided, use source text as the primary source and images only to resolve ambiguity.

        Image count: {{imageCount}}

        Source text:
        {{sourceText}}
        """;

    private static string BuildPrompt(string text, string targetLanguage, int imageCount)
    {
        (string, string?)[] values =
        [
            ("targetLanguage", targetLanguage),
            ("imageCount", imageCount.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            ("sourceText", string.IsNullOrWhiteSpace(text) ? "(none)" : text.Trim()),
        ];
        if (File.Exists(AiPromptRegistry.ResolvePath(TranslationPromptPath)))
            return AiPromptRegistry.Render(TranslationPromptPath, values);

        string prompt = FallbackTranslationPrompt;
        foreach ((string key, string? value) in values)
            prompt = prompt.Replace("{{" + key + "}}", value ?? "", StringComparison.Ordinal);
        return prompt;
    }
    private static JsonObject BuildTranslationJsonSchema()
        => new()
        {
            ["type"] = "json_schema",
            ["name"] = "translation_result",
            ["strict"] = true,
            ["schema"] = new JsonObject
            {
                ["type"] = "object",
                ["additionalProperties"] = false,
                ["properties"] = new JsonObject
                {
                    ["sourceText"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["description"] = "Exact AI OCR transcription or provided source text."
                    },
                    ["translation"] = new JsonObject
                    {
                        ["type"] = "string",
                        ["description"] = "Translation only."
                    }
                },
                ["required"] = new JsonArray
                {
                    JsonValue.Create("sourceText"),
                    JsonValue.Create("translation")
                }
            }
        };

    private static CodexTokenUsage? ExtractUsage(string raw)
    {
        using var doc = JsonDocument.Parse(raw);
        if (!doc.RootElement.TryGetProperty("usage", out var usage) || usage.ValueKind != JsonValueKind.Object)
            return null;

        int? input = TryGetInt(usage, "input_tokens");
        int? output = TryGetInt(usage, "output_tokens");
        int? total = TryGetInt(usage, "total_tokens");
        if (total is null && input is not null && output is not null)
            total = input.Value + output.Value;

        return input is null && output is null && total is null
            ? null
            : new CodexTokenUsage(input, output, total);
    }

    private static int? TryGetInt(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value) || value.ValueKind != JsonValueKind.Number)
            return null;
        return value.TryGetInt32(out int result) ? result : null;
    }

    private static string ExtractOutputText(string raw)
    {
        using var doc = JsonDocument.Parse(raw);
        var root = doc.RootElement;
        if (root.TryGetProperty("output_text", out var outputText))
            return outputText.GetString() ?? "";

        var sb = new StringBuilder();
        if (root.TryGetProperty("output", out var output) && output.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in output.EnumerateArray())
            {
                if (!item.TryGetProperty("content", out var content) || content.ValueKind != JsonValueKind.Array)
                    continue;

                foreach (var part in content.EnumerateArray())
                {
                    if (part.TryGetProperty("type", out var type)
                        && type.GetString() == "output_text"
                        && part.TryGetProperty("text", out var text))
                    {
                        sb.Append(text.GetString());
                    }
                }
            }
        }
        return sb.ToString();
    }

    private static CodexTranslationResult ParseTranslationResult(string output)
    {
        output = CleanJsonText(output);
        try
        {
            using var doc = JsonDocument.Parse(output);
            var root = doc.RootElement;
            string source = root.TryGetProperty("sourceText", out var s) ? s.GetString() ?? "" : "";
            string translation = root.TryGetProperty("translation", out var t) ? t.GetString() ?? "" : "";
            return new CodexTranslationResult(source.Trim(), translation.Trim(), null);
        }
        catch
        {
            return new CodexTranslationResult("", output.Trim(), null);
        }
    }

    private static string CleanJsonText(string text)
    {
        text = (text ?? "").Trim();
        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            int firstNewline = text.IndexOf('\n');
            if (firstNewline >= 0) text = text[(firstNewline + 1)..];
            int fence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (fence >= 0) text = text[..fence];
        }
        return text.Trim();
    }

    private static string FirstNonBlank(params string?[] values)
        => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? "";
}
