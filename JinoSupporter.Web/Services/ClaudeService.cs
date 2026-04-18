using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;

namespace JinoSupporter.Web.Services;

public sealed class ClaudeService
{
    private const string DefaultModel = "claude-haiku-4-5-20251001";
    // Claude's 5 MB limit applies to the base64-encoded payload; base64 grows ~33% over raw bytes.
    // Target 3.5 MB of raw bytes → ~4.67 MB base64, comfortably under the 5 MB cap.
    private const long   MaxImageBytes = 3_500_000;

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

    /// <summary>Compresses via <see cref="ImageCompressor"/> (quality-first, near-lossless).</summary>
    private static string ResizeImageBase64(string base64, string mediaType)
        => ImageCompressor.CompressIfLarge(base64, mediaType).Base64;

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

    // ── Normalize from images ─────────────────────────────────────────────────

    public async Task<NormalizeResult> NormalizeFromImagesAsync(
        List<(string MediaType, string Base64)> images,
        string datasetName,
        string productType,
        string testDate,
        CancellationToken ct = default)
    {
        if (!IsConfigured)
            throw new InvalidOperationException("Claude API key is not configured.");

        var blocks = new JsonArray();
        foreach ((string mediaType, string base64) in images)
        {
            string resized  = ResizeImageBase64(base64, mediaType);
            string detected = DetectMediaType(resized, mediaType);
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

        string prompt = $$"""
            You are a manufacturing quality data extraction specialist.

            Context:
            - Dataset: {{datasetName}}
            - Default product type: {{productType}}
            - Default test date: {{testDate}}

            The attached image(s) are screenshots of Excel manufacturing inspection reports.

            ══ STEP 1: IDENTIFY ALL TABLES AND SECTIONS ══
            A report may contain MULTIPLE TABLES (e.g. "RESULT CHECK PROCESS AI COIL", "RESULT CHECK PROCESS SPOT WELDING").
            A table may also contain MULTIPLE PRODUCT SECTIONS (e.g. TIA-338L rows vs TIA-338R rows in different merged-cell groups).

            For each table:
            → Read the TABLE NAME/HEADER → store in variableDetail for every row belonging to that table.
              Example: rows from "RESULT CHECK PROCESS AI COIL" get variableDetail = "AI COIL"
                       rows from "RESULT CHECK PROCESS SPOT WELDING" get variableDetail = "SPOT WELDING"
              This is CRITICAL — it is the only way to distinguish rows with the same variable name across tables.

            ══ STEP 2: READ MERGED CELLS CORRECTLY ══
            Merged cells span multiple rows. For each row group/section:
            - Model/product column → productType (if "{{productType}}" was given and image shows same, keep it)
            - Date column → testDate (YYYY-MM-DD; "12-Feb" with no year → "{{testDate}}"[..4] + "-02-12")
            - Line column → line (e.g. "C1120R", "C2-2B", "C2-1B")
            CRITICAL: Each product section has its OWN line. NEVER carry over a line from a different section.

            ══ STEP 3: DETERMINE variableGroup SEMANTICS ══
            variableGroup must reflect the ACTUAL comparison relationship:

            Row label contains "Before Ass'y" / "After Ass'y"  → "before" / "after"
            Row is test/modified condition (e.g. "Test ...", "New lot")  → "test" / "new_lot"
            Row is baseline/normal condition (e.g. "Normal ...", "Old lot")  → "normal" / "old_lot"
            Row label is just a Worker name with no phase  → leave variableGroup blank ""
            Row has "retrained" label  → variableGroup = "after", intervention = "retrained"

            ══ STEP 4: READ EACH DATA ROW ══
            For every non-total data row:
            - variable: primary identifier from Type/Worker/Condition column
            - variableDetail: TABLE NAME of the table this row belongs to (see STEP 1)
            - variableGroup: see STEP 3
            - intervention: "retrained" if row contains "retrained" or "재교육" (any language), else ""
            - inputQty: INTEGER from Input column — read every digit carefully
            - okQty: INTEGER from OK column
            - ngTotal: INTEGER from Total NG / Q'ty NG column
            - ngRate: FLOAT from NG rate column (strip %, e.g. "33.3%" → 33.3)
            - checkType: "process" / "function" / "visual_inspection" based on table context

            ══ STEP 5: WIDE → LONG TRANSFORM (DEDUPLICATED) ══
            Count how many defect columns have NON-ZERO values in each physical row, then apply:

            CASE A — 0 non-zero defect columns (clean row):
              → emit EXACTLY 1 entry with defectType="", defectCount=0,
                ngTotal/ngRate = row totals (may be 0).

            CASE B — EXACTLY 1 non-zero defect column (most common):
              → emit EXACTLY 1 entry MERGED:
                defectType = column label, defectCount = that cell value,
                ngTotal/ngRate = row totals, InputQty/OkQty = row values.
              → DO NOT also emit a separate "total" entry. This avoids duplication.

            CASE C — 2+ non-zero defect columns (multi-defect row):
              → emit 1 "total" entry (defectType="", defectCount=0,
                ngTotal/ngRate = row totals) — needed to preserve the aggregate
              → PLUS one entry per non-zero defect column
                (defectType = column label, defectCount = cell value).
              → All entries share the same InputQty/OkQty/ngTotal/ngRate.

            Zero-count defect columns are ALWAYS omitted — never create zero-count entries.

            ══ STEP 6: SKIP AGGREGATE ROWS ══
            Skip rows labeled "Total", "Grand Total", "Sum", sub-total rows.
            Extract only individual data rows.

            ══ STEP 7: DEFECT CATEGORY MAPPING ══
            assembly_defect    → VP+CD separate, glue, clamp, bond, coil separate
            cosmetic_defect    → damage, particle, scratch, burn, defrom/deform
            function_spl       → SPL NG, audiobus SPL
            function_thd       → THD NG
            function_hearing   → noise, hearing, audiobus
            wire_defect        → wire offset, wire forming, wire cutting, wire clamp, wire pad offset, solder weak
            magnetic_defect    → Gauss low/NG
            rear_visual_damage → rear damage position N
            other              → anything else not clearly listed above

            ══ STEP 8: TAG EXTRACTION — purpose-first, NOT column dump ══
            You must UNDERSTAND the report's intent, not list every value.
            Produce 4–10 high-signal tags. Quality over quantity.

            LANGUAGE RULE: **All tags MUST be in English only.**
            Translate any Korean terms from the source into concise English equivalents.
            Do NOT output Korean characters in tags under any circumstances.

            REQUIRED tags (include each if determinable from the report):

            1) **Main purpose (1 keyword)** — the single-word essence of why this report exists.
               Pick the closest English tag (lowercase, hyphen-free, short):
                 "lot-comparison", "process-improvement", "worker-evaluation",
                 "training-effect", "root-cause", "variation-analysis",
                 "new-lot-validation", "condition-optimization", "mold-comparison"
               Base it on the Purpose/Objective section AND the overall comparison
               structure — not on column names.

            2) **Review type(s)** — pick from actual sections seen (English only):
                 "process-inspection", "function-inspection",
                 "visual-inspection", "reliability-inspection"
               If multiple check types appear, include each.

            3) **Product/model** — as it appears: e.g. "TIU-C11-20", "BRS-161016", "TIA-338".
               Product codes are already English — keep as-is.

            OPTIONAL tags (include only if they add signal, English only):

            4) **Process or sub-process name** — the assembly/manufacturing step being
               evaluated, drawn from section headers or the Purpose text (NOT column headers).
               Translate any Korean to English:
                 e.g. "UV-drying" (not "UV 건조"),
                      "FPCB-assembly" (not "FPCB 조립"),
                      "VP-assembly", "AI-coil-process"

            5) **Key comparison variable** — what is being varied. English only:
                 e.g. "drying-time" (not "건조 시간"),
                      "mold-number" (not "Mold 번호"),
                      "lot-date", "worker-variance"

            6) **Intervention or action** — English only:
                 e.g. "retraining" (not "재교육"),
                      "lot-change", "condition-change"

            DO NOT output:
            - ❌ Any Korean characters (translate to English).
            - ❌ Defect/NG column names: "SPL", "Noise", "Touch", "THD", "Gauss low",
              "wire offset", "VP+CD separate", "Rear damage position 5", etc.
              These live in defectType/defectCategory fields already — DO NOT duplicate.
            - ❌ Specific numbers / percentages / counts: "3.6%", "477G".
            - ❌ Every row label verbatim: "Test level 1", "Test Dry UV Yoke 1 min",
              "Normal (AWF#1)". These are measurement rows, not tags.
            - ❌ Category enum values: "function_spl", "wire_defect", "assembly_defect".
            - ❌ Generic English words copied from table headers: "Input", "OK", "Total NG".
            - ❌ File/image metadata, dates, line codes.

            De-duplicate aggressively. Prefer semantic intent over surface wording.

            ══ OUTPUT ══
            Return ONLY valid JSON — no markdown fences, no extra text:
            {
              "measurements": [
                {
                  "productType": "",
                  "testDate": "",
                  "line": "",
                  "checkType": "",
                  "variable": "",
                  "variableDetail": "",
                  "variableGroup": "",
                  "intervention": "",
                  "inputQty": 0,
                  "okQty": 0,
                  "ngTotal": 0,
                  "ngRate": 0.0,
                  "defectCategory": "",
                  "defectType": "",
                  "defectCount": 0
                }
              ],
              "summary": "1–2 sentence description",
              "keyFindings": "Key findings (use \\n between bullet points)",
              "tags": ["keyword1", "keyword2", "..."]
            }
            """;

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        string raw = await CallWithContentAsync(blocks, 8192, ct, "claude-sonnet-4-6");
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
            return new NormalizeResult { Summary = "JSON parse error — check Claude response." };
        }
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

        string prompt = $$"""
            You are a manufacturing data classifier.

            Dataset name: {{datasetName}}
            Data context (table names + column labels + sample values):
            {{dataContext}}

            Tags already used in the database: {{existing}}

            Task: produce 3–8 concise tags that best describe this dataset.
            Rules:
            1. If any already-used tag is semantically equivalent or very similar to what you would suggest, use THAT EXACT EXISTING TAG verbatim.
            2. Only introduce a brand-new tag when no existing tag covers the concept.
            3. Each tag: 1–3 words, Title Case, English.
            4. Return ONLY a JSON array of strings. No explanation, no code fences.

            Example output: ["Wire Cutting", "Quality Control", "2024", "Defect Analysis"]
            """;

        string result = await CallAsync(prompt, 512, ct);
        return ParseJsonArray<string>(result);
    }

    // ── Ask AI from registered reports ────────────────────────────────────────

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
                                              CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(datasetsContext))
        {
            return new AskAiResult
            {
                Overall    = "등록된 리포트가 없어 답변할 수 없습니다. 먼저 Input Data 에서 리포트를 저장해주세요.",
                PerDataset = []
            };
        }

        string prompt = $$"""
            You are a manufacturing quality improvement assistant.

            A user has asked a question about a production problem. Answer it USING ONLY the information found in the registered dataset reports below.

            ══ STRICT RULES ══
            1. Do NOT use external/general knowledge. Only use facts present in the reports below.
            2. If no registered report contains relevant information, set "overall" to a short Korean notice that no relevant data was found, and return an empty "perDataset" array. Do not invent an answer.
            3. Produce ONE entry in "perDataset" for EVERY dataset that genuinely contributes to the answer. Copy "datasetName" VERBATIM from the "Dataset:" header in the context (full string, including numeric prefixes and spaces).
            4. In each per-dataset "answer": explain in Korean (한국어) what this SPECIFIC dataset shows and how it addresses the user's question. Cite concrete values from that dataset only (NG rate, defect type, product type, date, specific findings). 2–5 sentences is ideal.
            5. Do NOT include datasets that are irrelevant to the question.
            6. In "overall": give a 2–3 sentence Korean synthesis across the per-dataset findings — top recommendations in priority order. If there is only one relevant dataset, you may leave "overall" empty.
            7. Return ONLY valid JSON — no markdown fences, no extra commentary.

            ══ OUTPUT JSON SCHEMA ══
            {
              "overall": "2–3 sentence Korean overall recommendation across all datasets (may be empty).",
              "perDataset": [
                {
                  "datasetName": "<verbatim Dataset name>",
                  "answer": "Korean, dataset-specific answer with concrete numbers from this dataset."
                }
              ]
            }

            ══ USER QUESTION ══
            {{question}}

            ══ REGISTERED DATASET REPORTS ══
            {{datasetsContext}}
            """;

        string raw = await CallAsync(prompt, 4096, ct);
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
                ?? new AskAiResult { Overall = "(응답 파싱 실패)" };
        }
        catch
        {
            return new AskAiResult { Overall = raw };
        }
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
