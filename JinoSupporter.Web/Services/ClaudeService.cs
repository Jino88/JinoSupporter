using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services;

public sealed class ClaudeService
{
    private const string DefaultModel = "claude-haiku-4-5-20251001";
    // Claude's 5 MB limit applies to the base64-encoded payload; base64 grows ~33% over raw bytes.
    // Target 3.5 MB of raw bytes ??~4.67 MB base64, comfortably under the 5 MB cap.
    private const long   MaxImageBytes = 3_500_000;

    private readonly HttpClient _http;
    private readonly string     _apiKey;
    private readonly AiProviderSettingsService _providerSettings;

    public ClaudeService(HttpClient http, IConfiguration config, WebRepository repo, AiProviderSettingsService providerSettings)
    {
        _http = http;
        _providerSettings = providerSettings;
        // Priority: DB ??WpfSettingsReader (workhost-settings.json) ??appsettings.json
        string? fromDb  = repo.GetSetting("Claude:ApiKey");
        string? fromWpf = WpfSettingsReader.TryGetClaudeApiKey();
        string? fromCfg = config["Claude:ApiKey"];
        _apiKey = fromDb ?? fromWpf ?? fromCfg ?? string.Empty;
    }

    public bool IsConfigured => _providerSettings.ClaudeApiEnabled && !string.IsNullOrWhiteSpace(_apiKey);

    // ???? Core calls ????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    public async Task<string> CallAsync(string prompt,
                                        int maxTokens = 8192,
                                        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        var body = new JsonObject
        {
            ["model"]      = DefaultModel,
            ["max_tokens"] = maxTokens,
            ["messages"]   = new JsonArray
            {
                new JsonObject { ["role"] = "user", ["content"] = prompt }
            }
        };

        return await SendAsync(body, ct);
    }

    /// <summary>Sends a multi-part message (text + image blocks) using JsonNode to avoid List&lt;object&gt; serialization issues.</summary>
    private async Task<string> CallWithContentAsync(JsonArray contentBlocks,
                                                    int maxTokens = 8192,
                                                    CancellationToken ct = default,
                                                    string? model = null)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        var body = new JsonObject
        {
            ["model"]      = model ?? DefaultModel,
            ["max_tokens"] = maxTokens,
            ["messages"]   = new JsonArray
            {
                new JsonObject { ["role"] = "user", ["content"] = contentBlocks }
            }
        };

        return await SendAsync(body, ct);
    }

    private async Task<string> SendAsync(JsonObject body, CancellationToken ct)
    {
        string bodyJson = body.ToJsonString();
        using var request = new HttpRequestMessage(HttpMethod.Post, "messages")
        {
            Content = new StringContent(bodyJson, Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key",         _apiKey);
        request.Headers.Add("anthropic-version",  "2023-06-01");

        using HttpResponseMessage response = await _http.SendAsync(request, ct);
        string raw = await response.Content.ReadAsStringAsync(ct);

        if (!response.IsSuccessStatusCode)
        {
            // Surface payload size in the error so we can tell oversize-request from
            // other 4xx/5xx causes.
            double bodyMb = bodyJson.Length / 1024.0 / 1024.0;
            throw new HttpRequestException(
                $"Claude API error {(int)response.StatusCode} (request body {bodyMb:F1} MB): {raw}");
        }

        using JsonDocument doc = JsonDocument.Parse(raw);
        return doc.RootElement
                  .GetProperty("content")[0]
                  .GetProperty("text")
                  .GetString() ?? string.Empty;
    }

    // ???? Extract tables ????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    /// <summary>
    /// Extract tables from tab-separated text and/or images.
    /// When multiple images are provided without text, each image gets its own API call ??one table per image.
    /// </summary>
    public async Task<List<ExtractedTable>> ExtractTablesAsync(
        string rawData,
        List<(string MediaType, string Base64)>? images = null,
        CancellationToken ct = default)
    {
        bool imageOnly = images is { Count: > 0 } && string.IsNullOrWhiteSpace(rawData);
        bool perImage  = imageOnly && images!.Count > 1;

        if (perImage)
        {
            // Each image ??one call ??one table
            var all = new List<ExtractedTable>();
            int idx = 1;
            foreach ((string mediaType, string base64) in images!)
            {
                string resizedBase64 = ResizeImageBase64(base64, mediaType);
                string detectedType  = DetectMediaType(resizedBase64, mediaType);
                var blocks = BuildImageOnlyBlocks([(detectedType, resizedBase64)]);
                string result = await CallWithContentAsync(blocks, 8192, ct);
                List<ExtractedTable> tables = ParseJsonArray<ExtractedTable>(result);
                FillMergedCells(tables);
                foreach (ExtractedTable t in tables)
                {
                    if (string.IsNullOrWhiteSpace(t.TableName))
                        t.TableName = $"Image {idx}";
                    all.Add(t);
                }
                idx++;
            }
            return all;
        }

        if (imageOnly)
        {
            string resizedBase64 = ResizeImageBase64(images![0].Base64, images[0].MediaType);
            string detectedType  = DetectMediaType(resizedBase64, images[0].MediaType);
            var blocks = BuildImageOnlyBlocks([(detectedType, resizedBase64)]);
            string result = await CallWithContentAsync(blocks, 8192, ct);
            List<ExtractedTable> tables = ParseJsonArray<ExtractedTable>(result);
            FillMergedCells(tables);
            return tables;
        }

        if (images is { Count: > 0 })
        {
            // Text + images combined ??use per-image budget so total stays under limit
            var blocks = new JsonArray();
            foreach ((string resized, string detectedType) in ResizeImagesForBatch(images))
            {
                blocks.Add(new JsonObject
                {
                    ["type"]   = "image",
                    ["source"] = new JsonObject
                    {
                        ["type"]       = "base64",
                        ["media_type"] = detectedType,
                        ["data"]       = resized
                    }
                });
            }
            blocks.Add(new JsonObject { ["type"] = "text", ["text"] = BuildTextPrompt(rawData) });
            string result = await CallWithContentAsync(blocks, 8192, ct);
            List<ExtractedTable> tables = ParseJsonArray<ExtractedTable>(result);
            FillMergedCells(tables);
            return tables;
        }

        // Text only
        {
            string result = await CallAsync(BuildTextPrompt(rawData), 8192, ct);
            List<ExtractedTable> tables = ParseJsonArray<ExtractedTable>(result);
            FillMergedCells(tables);
            return tables;
        }
    }

    private static string BuildTextPrompt(string rawData)
    {
        string limited = rawData.Length > 40000 ? rawData[..40000] + "\n...(truncated)" : rawData;
        return AiPromptRegistry.Render("claude/text-table-parser.md", ("rawData", limited));
    }

    private static JsonArray BuildImageOnlyBlocks(List<(string MediaType, string Base64)> images)
    {
        string imagePrompt = AiPromptRegistry.Read("claude/image-table-parser.md");
        var blocks = new JsonArray();
        foreach ((string mediaType, string base64) in images)
        {
            blocks.Add(new JsonObject
            {
                ["type"]   = "image",
                ["source"] = new JsonObject
                {
                    ["type"]       = "base64",
                    ["media_type"] = DetectMediaType(base64, mediaType),
                    ["data"]       = base64
                }
            });
        }
        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = imagePrompt });
        return blocks;
    }

    // ???? Merged cell post-processing ??????????????????????????????????????????????????????????????????????????????????????

    private static void FillMergedCells(List<ExtractedTable> tables)
    {
        foreach (ExtractedTable table in tables)
        {
            if (table.Rows.Count == 0) continue;
            foreach (ColumnDef col in table.Columns)
            {
                string field = col.Field;

                // Pass 1: fill-UP ??first non-empty value propagated backward to leading empty rows
                string? firstVal = null;
                int firstIdx = -1;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (table.Rows[i].TryGetValue(field, out string? v) && !string.IsNullOrWhiteSpace(v))
                    { firstVal = v; firstIdx = i; break; }
                }
                if (firstVal is not null && firstIdx > 0)
                    for (int i = 0; i < firstIdx; i++)
                        table.Rows[i][field] = firstVal;

                // Pass 2: fill-DOWN
                string? last = null;
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    if (table.Rows[i].TryGetValue(field, out string? v) && !string.IsNullOrWhiteSpace(v))
                        last = v;
                    else if (last is not null)
                        table.Rows[i][field] = last;
                }
            }
        }
    }

    // ???? Image resize ????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    /// <summary>
    /// Detects the actual image format from the first few bytes of the data,
    /// returning one of the four media types Claude accepts.
    /// Falls back to the declared mediaType if detection is inconclusive.
    /// </summary>
    private static string DetectMediaType(string base64, string declaredMediaType)
    {
        try
        {
            // Decode just the first 12 bytes
            int needed = Math.Min(16, base64.Length / 4 * 3);
            byte[] header = Convert.FromBase64String(base64[..Math.Min(base64.Length, 20)]);

            if (header.Length >= 4 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
                return "image/png";
            if (header.Length >= 3 && header[0] == 0xFF && header[1] == 0xD8 && header[2] == 0xFF)
                return "image/jpeg";
            if (header.Length >= 6 && header[0] == 0x47 && header[1] == 0x49 && header[2] == 0x46)
                return "image/gif";
            if (header.Length >= 12 && header[0] == 0x52 && header[1] == 0x49 && header[2] == 0x46 && header[3] == 0x46
                && header[8] == 0x57 && header[9] == 0x45 && header[10] == 0x42 && header[11] == 0x50)
                return "image/webp";
        }
        catch { /* fall through */ }

        // Fall back to declared type, normalizing known aliases
        return declaredMediaType.ToLowerInvariant() switch
        {
            "image/jpg"  => "image/jpeg",
            "image/jpeg" => "image/jpeg",
            "image/png"  => "image/png",
            "image/gif"  => "image/gif",
            "image/webp" => "image/webp",
            _            => "image/png"
        };
    }

    private static string NormalizeMediaType(string mediaType) =>
        mediaType.ToLowerInvariant() switch
        {
            "image/jpg"  => "image/jpeg",
            "image/jpeg" => "image/jpeg",
            "image/png"  => "image/png",
            "image/gif"  => "image/gif",
            "image/webp" => "image/webp",
            _            => "image/png"
        };

    /// <summary>Single-image compression with default 3.5 MB target.</summary>
    private static string ResizeImageBase64(string base64, string mediaType)
        => ImageCompressor.CompressIfLarge(base64, mediaType).Base64;

    /// <summary>
    /// Multi-image compression ??divides a TOTAL budget across images so the
    /// aggregate request body stays under Anthropic's vision request size limit.
    /// Returns (resized base64, detected media type) per image.
    /// </summary>
    private static List<(string Base64, string MediaType)> ResizeImagesForBatch(
        IReadOnlyList<(string MediaType, string Base64)> images)
    {
        long perImageBudget = ImageCompressor.BudgetPerImage(images.Count);
        var results = new List<(string, string)>(images.Count);
        foreach ((string mediaType, string base64) in images)
        {
            (string newB64, string newMedia) = ImageCompressor.CompressIfLarge(
                base64, mediaType, perImageBudget);
            results.Add((newB64, DetectMediaType(newB64, newMedia)));
        }
        return results;
    }

    // ???? Generate HTML report ????????????????????????????????????????????????????????????????????????????????????????????????????

    public async Task<string> GenerateReportAsync(
        string datasetName,
        string tablesSummary,
        List<(string MediaType, string Base64)>? images = null,
        CancellationToken ct = default)
    {
        string prompt = AiPromptRegistry.Render(
            "claude/html-report.md",
            ("datasetName", datasetName),
            ("tablesSummary", tablesSummary));

        if (images is { Count: > 0 })
        {
            var blocks = new JsonArray();
            foreach ((string mediaType, string base64) in images)
                blocks.Add(new JsonObject
                {
                    ["type"]   = "image",
                    ["source"] = new JsonObject
                    {
                        ["type"] = "base64", ["media_type"] = DetectMediaType(base64, mediaType), ["data"] = base64
                    }
                });
            blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });
            return await CallWithContentAsync(blocks, 8192, ct);
        }

        return await CallAsync(prompt, 8192, ct);
    }

    // ???? OCR: Extract structured text from images (cacheable) ????????????????????????????????????

    /// <summary>
    /// Runs Claude vision once to produce a structured MARKDOWN transcript of
    /// every table, text section, and metadata in the report. Intended to be
    /// cached so downstream measurement extraction can run in text-only mode
    /// (cheaper, faster, and more debuggable). Compound headers are preserved
    /// via explicit labelling (e.g. "NG Audiobus: SPL+RB" as one column).
    /// </summary>
    public async Task<string> ExtractStructuredTextAsync(
        List<(string MediaType, string Base64)> images,
        string datasetName,
        string productType,
        string testDate,
        CancellationToken ct = default,
        string? rawExcelText = null)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        var blocks = new JsonArray();
        foreach ((string resized, string detected) in ResizeImagesForBatch(images))
        {
            blocks.Add(new JsonObject
            {
                ["type"]   = "image",
                ["source"] = new JsonObject
                {
                    ["type"]       = "base64",
                    ["media_type"] = detected,
                    ["data"]       = resized
                }
            });
        }

        string rawExcelBlock = string.IsNullOrWhiteSpace(rawExcelText)
            ? ""
            : AiPromptRegistry.Render("claude/structured-transcript-raw-excel-block.md", ("rawExcelText", rawExcelText.Trim()));
        string prompt = AiPromptRegistry.Render(
            "claude/structured-transcript.md",
            ("datasetName", datasetName),
            ("productType", productType),
            ("testDate", testDate),
            ("rawExcelBlock", rawExcelBlock));

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        string raw = await CallWithContentAsync(blocks, 64000, ct, "claude-sonnet-4-6");
        return raw.Trim();
    }

    // ???? Normalize from PRE-EXTRACTED TEXT (text-only, cheap) ????????????????????????????????????

    /// <summary>
    /// Text-only measurement extraction from a pre-OCR'd markdown transcript.
    /// Uses the SAME output JSON schema as <see cref="NormalizeFromImagesAsync"/>
    /// but skips all vision cost and is typically faster + more deterministic
    /// because the column structure is already resolved.
    /// </summary>
    public async Task<NormalizeResult> NormalizeFromTextAsync(
        string extractedText,
        string datasetName,
        string productType,
        string testDate,
        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        string prompt = AiPromptRegistry.Render(
            "claude/normalize-from-text.md",
            ("datasetName", datasetName),
            ("productType", productType),
            ("testDate", testDate),
            ("extractedText", extractedText));

        var blocks = new JsonArray { new JsonObject { ["type"] = "text", ["text"] = prompt } };
        string raw = await CallWithContentAsync(blocks, 64000, ct, "claude-sonnet-4-6");
        raw = raw.Trim();
        if (raw.StartsWith("```"))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```")) raw = raw[..raw.LastIndexOf("```")];
        }
        int open  = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        if (open >= 0 && close > open) raw = raw[open..(close + 1)];
        try
        {
            return JsonSerializer.Deserialize<NormalizeResult>(raw.Trim(),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new NormalizeResult();
        }
        catch
        {
            return new NormalizeResult { Summary = "JSON parse error -- check Claude response (text-mode)." };
        }
    }

    // ???? Normalize from images ??????????????????????????????????????????????????????????????????????????????????????????????????

    public async Task<NormalizeResult> NormalizeFromImagesAsync(
        List<(string MediaType, string Base64)> images,
        string datasetName,
        string productType,
        string testDate,
        CancellationToken ct = default,
        string? rawText = null)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        var blocks = new JsonArray();
        foreach ((string resized, string detected) in ResizeImagesForBatch(images))
        {
            blocks.Add(new JsonObject
            {
                ["type"]   = "image",
                ["source"] = new JsonObject
                {
                    ["type"]       = "base64",
                    ["media_type"] = detected,
                    ["data"]       = resized
                }
            });
        }

        string rawTextBlock = string.IsNullOrWhiteSpace(rawText)
            ? ""
            : AiPromptRegistry.Render("claude/normalize-raw-text-block.md", ("rawText", rawText.Trim()));
        string prompt = AiPromptRegistry.Render(
            "claude/normalize-from-images.md",
            ("datasetName", datasetName),
            ("productType", productType),
            ("testDate", testDate),
            ("rawTextBlock", rawTextBlock));

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        string raw = await CallWithContentAsync(blocks, 64000, ct, "claude-sonnet-4-6");
        raw = raw.Trim();
        if (raw.StartsWith("```"))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```")) raw = raw[..raw.LastIndexOf("```")];
        }
        int open  = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        if (open >= 0 && close > open) raw = raw[open..(close + 1)];
        try
        {
            return JsonSerializer.Deserialize<NormalizeResult>(raw.Trim(),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new NormalizeResult();
        }
        catch
        {
            return new NormalizeResult { Summary = "JSON parse error -- check Claude response." };
        }
    }

    // ???? Extract tags ????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    public async Task<List<string>> ExtractTagsAsync(string datasetName,
                                                     string memo,
                                                     string dataPreview,
                                                     CancellationToken ct = default)
    {
        string prompt = AiPromptRegistry.Render(
            "claude/extract-tags.md",
            ("datasetName", datasetName),
            ("memo", memo),
            ("dataPreview", dataPreview));

        string result = await CallAsync(prompt, 512, ct);
        return ParseJsonArray<string>(result);
    }

    /// <summary>
    /// Extract purpose tags from the dataset context, normalising against existing DB tags so
    /// that semantically equivalent concepts always use the same label.
    /// </summary>
    public async Task<List<string>> ExtractPurposeTagsAsync(
        string           datasetName,
        string           dataContext,
        List<string>     existingTags,
        CancellationToken ct = default)
    {
        string existing = existingTags.Count > 0
            ? string.Join(", ", existingTags)
            : "(none yet)";

        string prompt = AiPromptRegistry.Render(
            "claude/extract-purpose-tags.md",
            ("datasetName", datasetName),
            ("dataContext", dataContext),
            ("existingTags", existing));

        string result = await CallAsync(prompt, 512, ct);
        return ParseJsonArray<string>(result);
    }

    // ???? Ask AI from registered reports ????????????????????????????????????????????????????????????????????????????????

    public sealed class AskAiPerDataset
    {
        [JsonPropertyName("datasetName")] public string DatasetName { get; set; } = "";
        [JsonPropertyName("answer")]      public string Answer      { get; set; } = "";
    }

    public sealed class AskAiResult
    {
        [JsonPropertyName("overall")]    public string                 Overall    { get; set; } = "";
        [JsonPropertyName("perDataset")] public List<AskAiPerDataset>  PerDataset { get; set; } = [];
    }

    /// <summary>
    /// Answers <paramref name="question"/> using ONLY the provided registered dataset contexts.
    /// Returns an overall recommendation plus a per-dataset answer for every dataset that
    /// genuinely informs the answer.
    /// </summary>
    public async Task<AskAiResult> AskAiAsync(string question,
                                              string datasetsContext,
                                              string answerLanguage = "English",
                                              CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(datasetsContext))
        {
            return new AskAiResult
            {
                Overall    = "No registered reports ??cannot answer. Please save a report from Input Data first.",
                PerDataset = []
            };
        }

        string lang = string.IsNullOrWhiteSpace(answerLanguage) ? "English" : answerLanguage.Trim();

        string prompt = AiPromptRegistry.Render(
            "claude/ask-ai.md",
            ("answerLanguage", lang),
            ("question", question),
            ("datasetsContext", datasetsContext));

        string raw = await CallAsync(prompt, 12000, ct);
        raw = raw.Trim();
        if (raw.StartsWith("```"))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```")) raw = raw[..raw.LastIndexOf("```")];
        }
        int open  = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        if (open >= 0 && close > open) raw = raw[open..(close + 1)];

        try
        {
            return JsonSerializer.Deserialize<AskAiResult>(raw.Trim(),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? new AskAiResult { Overall = "(Failed to parse response)" };
        }
        catch
        {
            return new AskAiResult { Overall = raw };
        }
    }

    public async Task<AskAiResult> TranslateAskAiResultAsync(
        AskAiResult source,
        string targetLanguage,
        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        string inputJson = JsonSerializer.Serialize(source, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        string prompt = AiPromptRegistry.Render(
            "claude/translate-ask-ai-result.md",
            ("targetLanguage", targetLanguage),
            ("inputJson", inputJson));

        string raw = await CallAsync(prompt, 12000, ct);
        raw = raw.Trim();
        if (raw.StartsWith("```"))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```")) raw = raw[..raw.LastIndexOf("```")];
        }

        int open = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        if (open >= 0 && close > open) raw = raw[open..(close + 1)];

        return JsonSerializer.Deserialize<AskAiResult>(raw.Trim(),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? new AskAiResult { Overall = source.Overall, PerDataset = source.PerDataset };
    }

    // ???? Translate analysis result (multi-field, one round-trip) ??????????????????????????????

    /// <summary>Translate the 7 narrative fields of a NormalizeResult into the
    /// target language in a single API call. Returns the original record's
    /// values for any field the model couldn't translate ??never null fields.
    /// Tags / measurements are intentionally NOT translated (tags are domain
    /// keywords that should stay searchable; measurements are
    /// numeric/categorical structured data).</summary>
    public async Task<DatasetSummaryTranslation> TranslateAnalysisAsync(
        NormalizeResult source, string targetLanguage, CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        // v2 input: headline + actions[].text + context.{process,stage,baselineReason}
        // are the only translate-eligible new fields. The verdict enum, evidence
        // rows (all numeric / unit / part-code values), and tags stay verbatim
        // and are NOT sent to the translator.
        var actionTexts = (source.Actions ?? new List<ActionItem>())
            .Select(a => a?.Text ?? "")
            .ToList();

        var inputJson = JsonSerializer.Serialize(new
        {
            // legacy 7 fields ??still translated for old-schema rows
            summary           = source.Summary           ?? "",
            keyFindings       = source.KeyFindings       ?? "",
            purpose           = source.Purpose           ?? "",
            testConditions    = source.TestConditions    ?? "",
            rootCause         = source.RootCause         ?? "",
            decision          = source.Decision          ?? "",
            recommendedAction = source.RecommendedAction ?? "",

            // v2 ??new fields
            headline       = source.Headline ?? "",
            actionTexts    = actionTexts,
            contextProcess = source.Context?.Process        ?? "",
            contextStage   = source.Context?.Stage          ?? "",
            contextBaseline= source.Context?.BaselineReason ?? "",
        });

        string scriptRules = BuildOutputScriptRules(targetLanguage);

        string prompt = AiPromptRegistry.Render(
            "claude/translate-analysis.md",
            ("targetLanguage", targetLanguage),
            ("scriptRules", scriptRules),
            ("inputJson", inputJson));

        var content = new JsonArray
        {
            new JsonObject { ["type"] = "text", ["text"] = prompt },
        };

        string raw = (await CallWithContentAsync(content, maxTokens: 4096, ct: ct)).Trim();
        // Strip ``` ... ``` fence and clip to outermost { ... }.
        if (raw.StartsWith("```"))
        {
            int nl = raw.IndexOf('\n');
            if (nl >= 0) raw = raw[(nl + 1)..];
            if (raw.TrimEnd().EndsWith("```")) raw = raw[..raw.LastIndexOf("```")];
        }
        int open  = raw.IndexOf('{');
        int close = raw.LastIndexOf('}');
        string clean = (open >= 0 && close > open) ? raw[open..(close + 1)] : raw;

        // Best-effort parse. On any failure, fall back to the original (English)
        // values so the dataset still saves cleanly with a degraded translation.
        try
        {
            using var doc = JsonDocument.Parse(clean);
            var root = doc.RootElement;
            string Get(string k) => root.TryGetProperty(k, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString() ?? "" : "";

            // Re-assemble translated Actions array from actionTexts[]. Pair
            // each translated text with the original priority / kind from
            // source.Actions (those are not translated).
            var translatedActions = new List<ActionItem>();
            if (root.TryGetProperty("actionTexts", out var arr) && arr.ValueKind == JsonValueKind.Array)
            {
                var origActions = source.Actions ?? new List<ActionItem>();
                int i = 0;
                foreach (var el in arr.EnumerateArray())
                {
                    string text = el.ValueKind == JsonValueKind.String ? (el.GetString() ?? "") : "";
                    var orig = i < origActions.Count ? origActions[i] : null;
                    translatedActions.Add(new ActionItem
                    {
                        Priority = orig?.Priority ?? (i + 1),
                        Kind     = orig?.Kind     ?? "action",
                        Text     = text,
                    });
                    i++;
                }
            }

            AnalysisContext? translatedContext = null;
            string cp = Get("contextProcess");
            string cs = Get("contextStage");
            string cb = Get("contextBaseline");
            if (!string.IsNullOrEmpty(cp) || !string.IsNullOrEmpty(cs) || !string.IsNullOrEmpty(cb))
                translatedContext = new AnalysisContext { Process = cp, Stage = cs, BaselineReason = cb };

            return new DatasetSummaryTranslation
            {
                Summary           = Get("summary"),
                KeyFindings       = Get("keyFindings"),
                Purpose           = Get("purpose"),
                TestConditions    = Get("testConditions"),
                RootCause         = Get("rootCause"),
                Decision          = Get("decision"),
                RecommendedAction = Get("recommendedAction"),
                Headline          = Get("headline"),
                Actions           = translatedActions,
                Context           = translatedContext,
            };
        }
        catch
        {
            // Fall back to the original (English) values so the dataset still
            // saves cleanly with a degraded translation.
            return new DatasetSummaryTranslation
            {
                Summary           = source.Summary           ?? "",
                KeyFindings       = source.KeyFindings       ?? "",
                Purpose           = source.Purpose           ?? "",
                TestConditions    = source.TestConditions    ?? "",
                RootCause         = source.RootCause         ?? "",
                Decision          = source.Decision          ?? "",
                RecommendedAction = source.RecommendedAction ?? "",
                Headline          = source.Headline          ?? "",
                Actions           = source.Actions ?? new List<ActionItem>(),
                Context           = source.Context,
            };
        }
    }

    // ???? Translate ??????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    public async Task<string> TranslateAsync(string text,
                                             string targetLanguage,
                                             List<(string MediaType, string Base64)>? images = null,
                                             CancellationToken ct = default)
    {
        bool hasText  = !string.IsNullOrWhiteSpace(text);
        bool hasImage = images is { Count: > 0 };

        // Claude vision sometimes misreads ambiguous characters in screenshots and
        // emits them verbatim ??e.g. dropping a Chinese hanzi into a Vietnamese line
        // ("Th嶺???h???ろ렩??g d癲놁옓???kh??V??). Lock the output script with explicit rules so
        // every emitted character has to belong to the target language's writing
        // system. Translate semantically, never copy through unreadable glyphs.
        string scriptRules = BuildOutputScriptRules(targetLanguage);

        string instruction = (hasText, hasImage) switch
        {
            (true, false) => AiPromptRegistry.Render(
                "claude/translate-text-only.md",
                ("targetLanguage", targetLanguage),
                ("scriptRules", scriptRules),
                ("text", text)),

            (false, true) => AiPromptRegistry.Render(
                "claude/translate-image-only.md",
                ("targetLanguage", targetLanguage),
                ("scriptRules", scriptRules)),

            (true, true) => AiPromptRegistry.Render(
                "claude/translate-text-and-image.md",
                ("targetLanguage", targetLanguage),
                ("scriptRules", scriptRules),
                ("text", text)),

            _ => string.Empty,
        };

        if (string.IsNullOrEmpty(instruction))
            throw new ArgumentException("Provide text, image(s), or both to translate.");

        if (!hasImage)
            return await CallAsync(instruction, 4096, ct);

        var blocks = new JsonArray();
        foreach ((string mediaType, string base64) in images!)
        {
            blocks.Add(new JsonObject
            {
                ["type"]   = "image",
                ["source"] = new JsonObject
                {
                    ["type"]       = "base64",
                    ["media_type"] = DetectMediaType(base64, mediaType),
                    ["data"]       = base64,
                }
            });
        }
        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = instruction });
        return await CallWithContentAsync(blocks, 4096, ct);
    }

    /// <summary>Per-target-language writing-system constraint inserted into the
    /// translate prompt. The goal is to stop Claude from copying through OCR
    /// misreads ??e.g. dropping a Chinese hanzi into a Vietnamese line because
    /// the source character was ambiguous in the screenshot.</summary>
    private static string BuildOutputScriptRules(string targetLanguage)
    {
        string scriptLine = targetLanguage switch
        {
            "Korean"                => "Korean uses Hangul (???) plus standard Latin punctuation/numerals. Do NOT include Chinese hanzi, Japanese kana, or any other non-Hangul script.",
            "English"               => "English uses the unaccented Latin alphabet (A-Z, a-z) plus standard punctuation/numerals. Do NOT include Chinese, Korean, Japanese, or any non-Latin characters.",
            "Japanese"              => "Japanese uses Hiragana, Katakana, and Kanji (CJK characters that are part of the standard Japanese writing system). Do NOT include Hangul or non-Japanese scripts.",
            "Chinese (Simplified)"  => "Chinese (Simplified) uses Simplified Han characters plus standard punctuation/numerals. Do NOT include Traditional-only characters, Hangul, kana, or other non-Chinese scripts.",
            "Chinese (Traditional)" => "Chinese (Traditional) uses Traditional Han characters plus standard punctuation/numerals. Do NOT include Simplified-only characters, Hangul, kana, or other non-Chinese scripts.",
            "Vietnamese"            => "Vietnamese uses the Latin alphabet with diacritics (??嶺?癲?嶺?癲???嶺?嶺?嶺???????etc.) plus standard punctuation/numerals. Do NOT include Chinese hanzi, Korean Hangul, Japanese kana, or any non-Latin character ??every glyph must be a valid Vietnamese letter, digit, or punctuation mark.",
            "Spanish"               => "Spanish uses the Latin alphabet with diacritics (嶺?嶺?嶺?嶺?嶺?嶺?嶺? plus ??嶺?and standard punctuation/numerals. Do NOT include any non-Latin character.",
            "French"                => "French uses the Latin alphabet with diacritics (??嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺?嶺???嶺? plus standard punctuation/numerals. Do NOT include any non-Latin character.",
            "German"                => "German uses the Latin alphabet plus 嶺?嶺?嶺???and standard punctuation/numerals. Do NOT include any non-Latin character.",
            _                       => $"Output must use only the standard writing system for {targetLanguage}. Do NOT mix in characters from other scripts (Chinese hanzi, Korean Hangul, Japanese kana, etc.).",
        };

        return AiPromptRegistry.Render(
            "claude/output-script-rules.md",
            ("scriptLine", scriptLine));
    }

    // ???? Helper ????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????????

    private static List<T> ParseJsonArray<T>(string raw)
    {
        raw = raw.Trim();
        if (raw.StartsWith("```"))
        {
            int firstNewline = raw.IndexOf('\n');
            if (firstNewline >= 0) raw = raw[(firstNewline + 1)..];
            if (raw.TrimEnd().EndsWith("```"))
                raw = raw[..raw.LastIndexOf("```")];
        }

        // Extract first [...] block (skip leading prose)
        int arrOpen  = raw.IndexOf('[');
        int arrClose = raw.LastIndexOf(']');
        if (arrOpen >= 0 && arrClose > arrOpen)
            raw = raw[arrOpen..(arrClose + 1)];

        raw = raw.Trim();
        try
        {
            return JsonSerializer.Deserialize<List<T>>(raw,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? [];
        }
        catch { return []; }
    }
}
