using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace JinoSupporter.Web.Services;

public sealed record CodexTokenUsage(int? InputTokens, int? OutputTokens, int? TotalTokens);

public sealed record CodexTranslationResult(string SourceText, string Translation, CodexTokenUsage? Usage);

public sealed class CodexApiService(
    HttpClient http,
    IConfiguration config,
    WebRepository repo,
    AiProviderSettingsService providers)
{
    public const string DefaultTranslateModel = "gemini-3.6-flash";

    public bool IsConfigured => providers.CodexApiEnabled && !string.IsNullOrWhiteSpace(ApiKey);

    public string Model => FirstNonBlank(
        repo.GetSetting("OpenAI:TranslateModel"),
        repo.GetSetting("Codex:TranslateModel"),
        Environment.GetEnvironmentVariable("OPENAI_TRANSLATE_MODEL"),
        config["OpenAI:TranslateModel"],
        config["Codex:TranslateModel"],
        DefaultTranslateModel);

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
            ["max_output_tokens"] = 4096,
            ["store"] = false
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, "responses")
        {
            Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
        };
        request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", ApiKey);

        using HttpResponseMessage response = await http.SendAsync(request, ct);
        string raw = await response.Content.ReadAsStringAsync(ct);
        if (!response.IsSuccessStatusCode)
            throw new InvalidOperationException($"Codex API error {(int)response.StatusCode}: {raw}");

        string output = ExtractOutputText(raw);
        if (string.IsNullOrWhiteSpace(output))
            throw new InvalidOperationException("Codex API returned an empty response.");

        CodexTranslationResult result = ParseTranslationResult(output);
        return result with { Usage = ExtractUsage(raw) };
    }

    private static string BuildPrompt(string text, string targetLanguage, int imageCount)
        => AiPromptRegistry.Render(
            "translation/openai-ocr-translation.md",
            ("targetLanguage", targetLanguage),
            ("imageCount", imageCount.ToString(System.Globalization.CultureInfo.InvariantCulture)),
            ("sourceText", string.IsNullOrWhiteSpace(text) ? "(none)" : text.Trim()));
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
