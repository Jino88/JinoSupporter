using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace JinoSupporter.Web.Services;

public sealed class ClaudeService
{
    private const string DefaultModel = "claude-haiku-4-5-20251001";
    private const long   MaxImageBytes = 4_500_000; // 4.5 MB safe margin under the 5 MB limit

    private readonly HttpClient _http;
    private readonly string     _apiKey;

    public ClaudeService(HttpClient http, IConfiguration config, WebRepository repo)
    {
        _http = http;
        // Priority: DB → WpfSettingsReader (workhost-settings.json) → appsettings.json
        string? fromDb  = repo.GetSetting("Claude:ApiKey");
        string? fromWpf = WpfSettingsReader.TryGetClaudeApiKey();
        string? fromCfg = config["Claude:ApiKey"];
        _apiKey = fromDb ?? fromWpf ?? fromCfg ?? string.Empty;
    }

    public bool IsConfigured => !string.IsNullOrWhiteSpace(_apiKey);

    // ── Core calls ────────────────────────────────────────────────────────────

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
                new JsonObject { ["role"] = "user", ["content"] = contentBlocks }
            }
        };

        return await SendAsync(body, ct);
    }

    private async Task<string> SendAsync(JsonObject body, CancellationToken ct)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, "messages")
        {
            Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-api-key",         _apiKey);
        request.Headers.Add("anthropic-version",  "2023-06-01");

        using HttpResponseMessage response = await _http.SendAsync(request, ct);
        string raw = await response.Content.ReadAsStringAsync(ct);

        if (!response.IsSuccessStatusCode)
            throw new HttpRequestException($"Claude API error {(int)response.StatusCode}: {raw}");

        using JsonDocument doc = JsonDocument.Parse(raw);
        return doc.RootElement
                  .GetProperty("content")[0]
                  .GetProperty("text")
                  .GetString() ?? string.Empty;
    }

    // ── Extract tables ────────────────────────────────────────────────────────

    /// <summary>
    /// Extract tables from tab-separated text and/or images.
    /// When multiple images are provided without text, each image gets its own API call → one table per image.
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
            // Each image → one call → one table
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
            // Text + images combined
            var blocks = new JsonArray();
            foreach ((string mediaType, string base64) in images)
            {
                string resized      = ResizeImageBase64(base64, mediaType);
                string detectedType = DetectMediaType(resized, mediaType);
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
        return
            "You are a manufacturing data parser. Parse the following tab-separated clipboard data from an Excel table.\n\n" +
            "STEP 1: For each line, split by the TAB character (\\t), assigning column indices 0, 1, 2, ... to each cell.\n\n" +
            "STEP 2: Identify header rows (typically the first 1-3 lines before numeric data rows).\n" +
            "- Within each header row, forward-fill empty cells left->right: an empty cell inherits the nearest non-empty label to its left.\n" +
            "- If there are multiple header rows, concatenate labels at the SAME column index (space-separated).\n" +
            "- Example: row 0 col 7 = \"NG AUDIOBUS\" (forward-filled), row 1 col 7 = \"SPL\" -> final label = \"NG AUDIOBUS SPL\".\n\n" +
            "STEP 3: Fill down merged cells in data rows.\n" +
            "- Excel merged cells paste as a value only in the first row; subsequent rows have the same column empty.\n" +
            "- For each data column, if a cell is empty, copy the most recent non-empty value from an earlier row.\n\n" +
            "STEP 4: Exclude rows where all cells are empty, or the row is a percentage sub-row, or a grand total/summary row.\n\n" +
            "STEP 5: Return ONLY a valid JSON array (no markdown fences, no explanation):\n" +
            "[\n  {\n    \"tableName\": \"descriptive name\",\n" +
            "    \"columns\": [{\"field\": \"f0\", \"label\": \"Column Label\"}, ...],\n" +
            "    \"rows\": [{\"f0\": \"value\", ...}, ...]\n  }\n]\n\n" +
            "CRITICAL: Use column index arithmetic only — never infer a column label from data values or neighboring columns.\n\n" +
            "DATA:\n" + limited;
    }

    private static JsonArray BuildImageOnlyBlocks(List<(string MediaType, string Base64)> images)
    {
        const string imagePrompt = """
            The attached image(s) are Excel sheet screenshots containing manufacturing inspection or production data.

            【Rules — must follow strictly】

            ▶ STEP 1. Merged cell handling — CRITICAL
               Excel merged cells show one value spanning multiple rows/columns visually.
               When unmerging, EVERY cell in the merged range must receive that value — including the FIRST row/column of the range.
               The value belongs to the TOP-LEFT cell; all other cells in the block copy it.
               - Horizontal merge: copy the value into EACH column it covers.
               - Vertical merge: copy the value into EACH row it covers, starting from the FIRST row.
                 e.g. 'Model A' visually spans rows 1-5 → rows 1, 2, 3, 4, 5 ALL get 'Model A' (including row 1)
               - Combined merge: fill EVERY cell in the entire block.
               DO NOT leave any cell empty that was part of a merged range.
               DO NOT skip the first row of a vertical merge.

            ▶ STEP 2. Multi-row headers
               - Apply STEP 1 merge-fill first on each header row independently.
               - Concatenate header rows top-to-bottom per column, omitting exact duplicate words.
                 e.g. col 7: row1='NG AUDIOBUS', row2='SPL' → label='NG AUDIOBUS SPL'

            ▶ STEP 3. Data rows
               - Apply STEP 1 merge-fill to every data cell (vertical merges across rows are common).
               - Include rows with actual measurements or row numbers.
               - Exclude total/subtotal/average/grand total/blank rows.

            ▶ STEP 4. Output — return JSON only (no ``` or other text)
            [
              {
                "tableName": "descriptive name",
                "columns": [{"field": "camelCaseEnglish", "label": "OriginalHeaderName"}],
                "rows": [{"field": "value", ...}, ...]
              }
            ]
               - field: English camelCase identifier
               - label: original header text as read from the image
               - all values are strings
            """;

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

    // ── Merged cell post-processing ───────────────────────────────────────────

    private static void FillMergedCells(List<ExtractedTable> tables)
    {
        foreach (ExtractedTable table in tables)
        {
            if (table.Rows.Count == 0) continue;
            foreach (ColumnDef col in table.Columns)
            {
                string field = col.Field;

                // Pass 1: fill-UP — first non-empty value propagated backward to leading empty rows
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

    // ── Image resize ──────────────────────────────────────────────────────────

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

    /// <summary>
    /// If the decoded byte size exceeds MaxImageBytes, scales the image down until it fits.
    /// Uses System.Drawing (available via Microsoft.AspNetCore.App on Windows) or falls back gracefully.
    /// </summary>
    private static string ResizeImageBase64(string base64, string mediaType)
    {
        byte[] data;
        try { data = Convert.FromBase64String(base64); }
        catch { return base64; }

        if (data.Length <= MaxImageBytes) return base64;

        try
        {
            // Estimate scale factor from size ratio (pixel area ~ file size)
            double scale = Math.Sqrt((double)MaxImageBytes / data.Length) * 0.85;

            using var inputMs = new MemoryStream(data);
            using var bmp = System.Drawing.Image.FromStream(inputMs);

            int newW = Math.Max(400, (int)(bmp.Width  * scale));
            int newH = Math.Max(300, (int)(bmp.Height * scale));

            // Try progressively smaller sizes
            for (int attempt = 0; attempt < 5; attempt++)
            {
                using var resized = new System.Drawing.Bitmap(newW, newH);
                using (var g = System.Drawing.Graphics.FromImage(resized))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.Clear(System.Drawing.Color.White);
                    g.DrawImage(bmp, 0, 0, newW, newH);
                }

                using var outMs = new MemoryStream();
                // Save as PNG for lossless; if still too big try JPEG
                resized.Save(outMs, System.Drawing.Imaging.ImageFormat.Png);
                if (outMs.Length <= MaxImageBytes)
                    return Convert.ToBase64String(outMs.ToArray());

                // Try JPEG
                using var jpgMs = new MemoryStream();
                var jpgEncoder = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                    .First(e => e.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                var encParams = new System.Drawing.Imaging.EncoderParameters(1);
                encParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality, 70L);
                resized.Save(jpgMs, jpgEncoder, encParams);
                if (jpgMs.Length <= MaxImageBytes)
                    return Convert.ToBase64String(jpgMs.ToArray());

                newW = (int)(newW * 0.7);
                newH = (int)(newH * 0.7);
            }
        }
        catch
        {
            // If resize fails, return original and let API reject if too large
        }

        return base64;
    }

    // ── Generate HTML report ──────────────────────────────────────────────────

    public async Task<string> GenerateReportAsync(
        string datasetName,
        string tablesSummary,
        List<(string MediaType, string Base64)>? images = null,
        CancellationToken ct = default)
    {
        string prompt = $$"""
            You are a manufacturing quality analyst. Based on the following dataset summary, generate a professional HTML report.

            Dataset: {{datasetName}}

            Data:
            {{tablesSummary}}

            Requirements:
            - Write a full, self-contained HTML document with embedded CSS (no external dependencies)
            - Include an executive summary at the top
            - Render each table as an HTML <table> with proper headers and styling
            - Add a brief analysis section after each table: notable trends, NG patterns, anything unusual
            - If reference images are provided, use them as visual context to enrich analysis
            - Use a clean, professional style (white background, readable fonts, subtle borders)
            - All text must be in English
            - Return ONLY the HTML document, no markdown fences, no explanation
            """;

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

    // ── Extract tags ──────────────────────────────────────────────────────────

    public async Task<List<string>> ExtractTagsAsync(string datasetName,
                                                     string memo,
                                                     string dataPreview,
                                                     CancellationToken ct = default)
    {
        string prompt = $$"""
            Extract 3-7 concise descriptive tags for categorizing this manufacturing dataset.
            Dataset name: {{datasetName}}
            Memo: {{memo}}
            Data preview (first lines):
            {{dataPreview}}

            Return ONLY a JSON array of strings, for example: ["AudioBus", "NG Analysis", "Q1 2024"]
            No explanation, no code fences.
            """;

        string result = await CallAsync(prompt, 512, ct);
        return ParseJsonArray<string>(result);
    }

    // ── Translate ─────────────────────────────────────────────────────────────

    public async Task<string> TranslateAsync(string text,
                                             string targetLanguage,
                                             CancellationToken ct = default)
    {
        string prompt = $$"""
            Translate the following text to {{targetLanguage}}.
            Return only the translation — no explanation, no original text, no commentary.

            {{text}}
            """;

        return await CallAsync(prompt, 4096, ct);
    }

    // ── Helper ────────────────────────────────────────────────────────────────

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
