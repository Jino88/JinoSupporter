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
            // Text + images combined — use per-image budget so total stays under limit
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

    /// <summary>Single-image compression with default 3.5 MB target.</summary>
    private static string ResizeImageBase64(string base64, string mediaType)
        => ImageCompressor.CompressIfLarge(base64, mediaType).Base64;

    /// <summary>
    /// Multi-image compression — divides a TOTAL budget across images so the
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

    // ── OCR: Extract structured text from images (cacheable) ──────────────────

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
        CancellationToken ct = default)
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

        string prompt = $$"""
            You are a data transcription specialist. Convert the attached manufacturing
            inspection report image(s) into a STRUCTURED MARKDOWN TRANSCRIPT that
            preserves every table, cell value, and section header with no inference
            or summarisation. A separate step will later parse the transcript into
            measurements; your job is ONLY accurate transcription.

            Context:
            - Dataset: {{datasetName}}
            - Default product type: {{productType}}
            - Default test date: {{testDate}}

            ══ OUTPUT FORMAT (markdown only, no JSON, no commentary) ══

            1) Start with a metadata block:
               ```
               # {{datasetName}}
               - Title: <report title as printed>
               - Date: <header Date cell>
               - Marker/Dept/Line: <whatever is printed>
               ```

            2) For each section "I. Purpose", "II. Content", "III. Result",
               "IV. Decision" (and any other top-level section), emit:
               ```
               ## <section name>
               <verbatim text content, line-by-line>
               ```

            3) For each TABLE in the report, emit:
               ```
               ### Table: <table heading if any, else "Untitled Table N">
               Columns: <pipe-separated LEAF column names, with parent prefix>
               Rows:
               | <cell1> | <cell2> | ... |
               | <cell1> | <cell2> | ... |
               ```

               LEAF COLUMN NAMING (critical for downstream parsing):
               - If a merged super-header groups sub-columns (e.g. "NG AUDIOBUS"
                 spans "SPL | SPL+RB | RB | No sound"), prefix EACH sub-column
                 with the parent: `NG Audiobus: SPL`, `NG Audiobus: SPL+RB`,
                 `NG Audiobus: RB`, `NG Audiobus: No sound`.
               - Compound labels that contain `+`/`&`/`/` are SINGLE columns; do
                 not split them.
               - Preserve Korean/non-English labels as-is.

               ROW TRANSCRIPTION:
               - Read ALL rows, including "Normal"/"Baseline" rows, total rows,
                 and rows with mostly zeros.
               - Preserve empty cells as empty (` `), zeros as `0`, and merged
                 cells by repeating the value on continuation rows.
               - Number rows exactly as they appear; do not reorder.

               PERCENT SUB-ROWS (continuation rows showing derived percentages):
               Manufacturing reports often put a "count row" and a "percent row"
               together — the percent row has BLANK identifier cells
               (No / Date / Model / Type / Input / OK) and contains only X.X%
               values derived from the counts above it.
               → Mark these sub-rows explicitly with a `(%)` flag at the start:
                 `| (%) |   |   |   |   |   |   | 0.0% | 0.0% | 6.7% | … |`
               The downstream parser will skip them as derived data.

               CONTINUATION ROWS WITH RAW DATA (rare):
               If a row has blank identifiers BUT contains additional RAW COUNT
               cells (not percentages) that clearly belong to the row above,
               prefix with `(cont)` so the parser can merge: `| (cont) | … |`.

            4) For each IMAGE / PHOTO panel inside the report (sample defect
               pictures, charts without numeric data), emit a single line:
               `![image] <brief caption if labelled, else "sample at row/col
               position N">`
               Do not hallucinate what the image contains.

            5) At the bottom, emit:
               ```
               ## Raw footnotes
               <any author comments / arrows / annotations verbatim>
               ```

            ══ STRICT RULES ══
            - Transcribe, do NOT summarise, interpret, or normalise.
            - Numbers exactly as printed: "14.9%", "US$2.93", "3,600".
            - Dates exactly as printed: "12-Feb", "2025-12-05".
            - When a value is illegible, write `[?]` — do not guess.
            - Column count in header MUST equal column count in every row.
            - No JSON. No code fences around the whole output. Output markdown directly.
            """;

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        string raw = await CallWithContentAsync(blocks, 16384, ct, "claude-sonnet-4-6");
        return raw.Trim();
    }

    // ── Normalize from PRE-EXTRACTED TEXT (text-only, cheap) ──────────────────

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

        string prompt = $$"""
            You are a manufacturing quality data extraction specialist. The input
            is ALREADY A STRUCTURED MARKDOWN TRANSCRIPT of a report (tables + text
            sections). Your job is to produce the normalised measurements JSON.

            Context:
            - Dataset: {{datasetName}}
            - Default product type: {{productType}}
            - Default test date: {{testDate}}

            All the rules below about column alignment, case handling, Normal rows,
            self-checks, etc. apply IDENTICALLY to text input — treat each
            markdown table as the source of truth.

            ══ CORE RULES (condensed) ══

            • LAYOUT CLASSIFICATION — classify each table into:
              (A) Standard / (B) Multi-stage funnel / (C) Aggregate-only /
              (D) Criterion-level / (E) Picture sample catalog / (F) Visual/waveform ref.
              Branch extraction rules accordingly (see below).

            • MULTI-STAGE FUNNEL: for each sub-stage row, use THAT stage's Input /
              NG count, not the row-level roll-up. variableDetail encodes the stage.

            • AGGREGATE-ONLY: emit ONE row with defectType="" — legitimate, not a bug.

            • CRITERION-LEVEL (Wire Moving / Frame Deform / …): ngTotal = MAX of
              per-criterion NGs, NOT Input-OK. Don't emit rows for the "OK" halves.

            • PICTURE SAMPLE (ppm column): put ppm/10000 in ngRate, leave ngTotal=0,
              defectCount=0. Don't create rows for individual sample photos.

            • VISUAL/WAVEFORM REF: at most ONE aggregate row summarising the test.
              Individual specimen IDs (OK #1, NG #2) are NOT measurements.

            • NORMAL ROWS: extract defect breakdowns with the SAME method as Test
              rows. Never leave a Normal row with ngTotal>0 and defectType="".

            • COMPOUND LABELS (`A+B`, `X&Y`): ONE column. The transcript already
              preserved these via `Parent: A+B` naming — use the label verbatim.

            • SKIP DERIVED ROWS (critical for text mode):
              - Rows prefixed `(%)` in the transcript are percent sub-rows of the
                count row ABOVE. DO NOT emit as a separate measurement.
              - Rows prefixed `(cont)` are continuation of the preceding row.
                Merge values into the previous row's entry, not a new row.
              - Rows labeled "Total" / "Grand Total" / "Sum" / sub-total rows — skip.

            • WIDE → LONG TRANSFORM:
              - CASE A (0 non-zero defect cols): 1 entry, defectType="", count=0.
              - CASE B (1 non-zero): 1 MERGED entry (defectType=col, count=cell).
              - CASE C (2+ non-zero): 1 aggregate entry + N per-defect entries.

            • variableGroup semantics: Normal→"normal", Test→"test",
              Before/After→"before"/"after", new lot→"new_lot", etc.

            • checkType: "process" / "function" / "visual_inspection" from table context.

            • defectCategory mapping (use only these enums):
              assembly_defect / cosmetic_defect / function_spl / function_thd /
              function_hearing / wire_defect / magnetic_defect / rear_visual_damage / other.

            ══ STRUCTURED CONTEXT FIELDS ══

            Extract 5 focused fields (empty string "" if not stated):
            - purpose           — goal of the test (from "I. Purpose" section)
            - testConditions    — what was varied
            - rootCause         — identified cause if concluded
            - decision          — final verdict (from "IV. Decision" section)
            - recommendedAction — concrete next step

            ══ TAGS ══
            Produce 4–10 English purpose-first tags (lowercase, hyphenated).
            Include product code, review type, main purpose, key comparison
            variable, intervention if present.

            ══ SELF-CHECK ══

            Before output, verify:
            (a) Every Total NG > 0 row has ≥1 defect entry when the table had
                per-defect columns.
            (b) Sum of positive defectCounts for a row ≥ ngTotal.
            (c) defectType strings match column labels from the transcript exactly.
            (d) No Normal rows left empty when defect columns are present.

            ══ INPUT TRANSCRIPT ══
            {{extractedText}}

            ══ OUTPUT (strict JSON only) ══
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
              "tags": ["keyword1", "keyword2", "..."],
              "purpose": "",
              "testConditions": "",
              "rootCause": "",
              "decision": "",
              "recommendedAction": ""
            }
            """;

        var blocks = new JsonArray { new JsonObject { ["type"] = "text", ["text"] = prompt } };
        string raw = await CallWithContentAsync(blocks, 16384, ct, "claude-sonnet-4-6");
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
            return new NormalizeResult { Summary = "JSON parse error — check Claude response (text-mode)." };
        }
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

        string prompt = $$"""
            You are a manufacturing quality data extraction specialist.

            Context:
            - Dataset: {{datasetName}}
            - Default product type: {{productType}}
            - Default test date: {{testDate}}

            The attached image(s) are screenshots of Excel manufacturing inspection reports.

            ══ STEP 0: CLASSIFY EACH TABLE'S LAYOUT ══
            Before extracting, classify each table into ONE of these layouts. The
            extraction rules BRANCH on this classification — misclassifying leaks
            bad data downstream.

            (A) STANDARD — Row has (Input, OK, Total NG, NG%) + per-defect count
                columns (SPL/RB/Noise/etc). Most common. Extract normally.

            (B) MULTI-STAGE FUNNEL — Row has row-level totals (Input, OK, Total
                NG, Total rate) PLUS repeated sub-groups like
                "Before Function | After Function | Visual Final" each with their
                OWN (Input, NG xxx, NG rate). The SAME unit is re-measured at each
                stage; sub-group Inputs DECREASE (200 → 179 → 173) as OK units
                carry forward.
                → RULE: For each sub-stage emit ONE entry where
                  InputQty = THAT stage's Input (not the row-level 200),
                  NgTotal  = THAT stage's NG count (not the row-level 27),
                  NgRate   = THAT stage's NG rate,
                  variableDetail encodes the stage name (e.g.
                  "VP damage: before function", "VP damage: after function").
                → NEVER copy the row-level Total NG/Input into every stage — those
                  are a rollup across stages, not per-stage values.

            (C) AGGREGATE-ONLY — Row has only (Input, OK, Total NG, NG rate) with
                NO per-defect breakdown columns. Example: "RA LINE" sub-tables with
                just Type/Input/OK/Total NG.
                → RULE: Emit ONE entry per row with defectType="", defectCount=0,
                  ngTotal from the row. This is LEGITIMATE — do NOT invent defect
                  subtypes. The data simply doesn't have a breakdown.

            (D) CRITERION-LEVEL EVALUATION — Row has multiple OK/NG sub-column
                PAIRS per criterion (e.g. "Wire Moving OK | Wire Moving NG" AND
                "Frame Deform OK | Frame Deform NG"). A unit can be overall-OK
                while having one criterion NG. Input=OK holds for hard rejects
                only; criterion NGs are advisory.
                → RULE: Emit one entry per criterion using its NG count as
                  defectCount. Set ngTotal = MAX of per-criterion NG counts (not
                  Input-OK which can legitimately be 0). Do NOT emit rows for the
                  "OK" half of each pair — that's not a defect.

            (E) PICTURE SAMPLE CATALOG — Table rows are defect TYPES (NG Damage,
                F-PCB separate…) with "NG Rate (ppm)" and example photos. No per-
                process counts.
                → RULE: Emit ONE row per NG Type. Put ppm value in ngRate
                  (divide by 10000 for %: 15383 ppm → 1.5383%). Leave inputQty,
                  okQty, ngTotal at 0. defectType = NG Type name,
                  defectCategory = mapped category, defectCount = 0.
                → Do NOT create rows for individual sample photos / column headers
                  of photo cells.

            (F) VISUAL / WAVEFORM REFERENCE — Table shows individual specimen IDs
                (OK #1…#11, NG #2, #4, #6) tied to photos or frequency graphs with
                diff annotations (e.g. "Diff: 12.39dB @ 9000Hz").
                → RULE: Do NOT emit per-specimen rows. Emit at most ONE aggregate
                  row summarising the test (e.g. "Frame of 11 units, 3 NG = 27.3%").
                  Individual specimen IDs are NOT measurements.

            If a single REPORT mixes layouts (e.g. a Standard table AND a RA-line
            Aggregate-only table), apply the appropriate rule to each table
            independently.

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

            CONTINUATION / PERCENT SUB-ROWS:
            Many reports print a "count row" with identifiers (No/Date/Model/Type/
            Input/OK/…) immediately followed by a sub-row with BLANK identifiers
            that shows derived PERCENTAGES of the same data. Example:
              | 1 | 9-Jan | Test… | 60 | 57 | 0 | 0 | 4 | 0 | 3 | 0 | 3 | 5.0% |
              |   |       |       |    |    |0.0%|0.0%|6.7%|0.0%|5.0%|0.0%|   |      |
            The percent row is DERIVED, NOT a separate measurement.
            → Treat these two lines as ONE logical data row. Extract from the
              COUNT row only. DO NOT emit a phantom measurement for the percent row.
            → If the OCR transcript marks it with `(%)` prefix, that's your signal.
            → Continuation rows (blank ids + raw counts continuing the row above)
              marked `(cont)` should be merged into the same logical row.

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
            - variableDetail: TABLE NAME of the table this row belongs to (see STEP 1).
              For MULTI-STAGE FUNNEL (layout B) append the stage name.
            - variableGroup: see STEP 3
            - intervention: "retrained" if row contains "retrained" or "재교육" (any language), else ""
            - inputQty: INTEGER from Input column — read every digit carefully.
              For FUNNEL: use the STAGE's Input, not the row-level roll-up.
            - okQty: INTEGER from OK column
            - ngTotal: INTEGER from Total NG / Q'ty NG column.
              For FUNNEL: use THAT stage's NG count, NOT the row-level Total NG.
              For CRITERION: use MAX of criterion NG counts.
              For PICTURE SAMPLE / VISUAL REF: set to 0 (no count available).
            - ngRate: FLOAT from NG rate column (strip %, e.g. "33.3%" → 33.3).
              For PICTURE SAMPLE ppm: divide by 10000 (15383 ppm → 1.5383).
            - checkType: "process" / "function" / "visual_inspection" based on table context

            ══ STEP 4.5: NO ROW CLASS IS EXEMPT FROM DEFECT EXTRACTION ══
            Every physical data row with ngTotal>0 must have its defect breakdown
            extracted if the table has per-defect columns — regardless of the row's
            label or semantic role. This includes (but is not limited to):
              - "Normal" / "Baseline" / "Reference" / "기존" baseline rows
              - "Total" / sub-total / summary rows (if they're the only row shown)
              - Rows with long descriptive names or non-English labels
              - Rows in "normal" VariableGroup just as much as "test" group
            NEVER leave a row with ngTotal>0 and empty defectType when the image
            clearly shows non-zero per-defect counts for that row. Selectively
            skipping certain row classes is a recurring extraction bug — extract
            ALL rows consistently with the same method.

            ══ STEP 5: WIDE → LONG TRANSFORM (DEDUPLICATED) ══

            ══ COLUMN ALIGNMENT — CRITICAL ══
            Defect columns sit UNDER merged section headers (e.g.
            "NG VP SEPARATE" spans 2 cols, "NG AUDIOBUS" spans 4 cols, "NG HEARING" spans 2 cols).
            Before reading any row value you MUST:
              1) Enumerate the leaf column headers strictly LEFT → RIGHT in the order they
                 visually appear, INCLUDING every column under every merged section.
                 Example for the header above: [VPSep:Noise, VPSep:Touch, Audio:SPL,
                 Audio:SPL+RB, Audio:RB, Audio:No sound, Hearing:Noise, Hearing:Touch].
              2) For each data row, match cell values to those leaf columns BY POSITION
                 ONE-TO-ONE. The i-th numeric cell in the row belongs to the i-th leaf column.
              3) NEVER skip a column. NEVER swap adjacent-section columns (e.g. do NOT label
                 an "Audio:RB" cell as "Audio:No sound", and do NOT label an "Audio:SPL" cell
                 as "VPSep:Noise"). If the image is ambiguous, favor the leftmost column of
                 the next section rather than guessing across sections.
              3a) COMPOUND vs SIMPLE LABELS — general rule, applies to ANY columns:
                  Some column headers are COMPOUND: a single label expressing the
                  INTERSECTION or COMBINATION of two concepts, written with a joining
                  character like `+`, `&`, `/`, `and`, `or`, `with`. A compound label
                  occupies exactly ONE cell — the same as any simple label.

                  RULE: Any header containing a joining marker (e.g. "A+B", "A&B",
                  "A/B", "A and B") is ONE column, NOT two. Do NOT split it across
                  two value cells. Do NOT conflate it with a neighbouring simple
                  label that shares a substring (e.g. a header "X+Y" is NOT the same
                  as adjacent "X" or "Y" alone — those are three distinct columns).

                  Instances seen in past reports (not exhaustive):
                    • Audio / function groups commonly use "SPL", "SPL+RB", "RB",
                      "No sound" side-by-side — four distinct columns.
                    • Function-check groups use "FRF", "FRF+SPL", "THD", "No sound" —
                      also four distinct columns.
                    • Hearing groups use "Noise", "Touch" — two simple columns.
                  Whatever labels actually appear, apply the same compound-vs-simple
                  rule: count leaf headers; never merge or split.

                  SANITY CHECK: Before emitting a row, COUNT the leaf headers you
                  enumerated in step (1) and COUNT the value cells in the row. They
                  MUST be equal. If unequal, you have mis-identified the header
                  structure — re-scan the image from left to right.
              4) If a row has a cell with a value but no leaf column maps to that position
                 (alignment appears broken), STOP and emit a defectType="__ALIGN_ERROR__"
                 entry for that row instead of guessing.
              5) The defectType string MUST be the EXACT leaf column label (including the
                 section prefix if the plain name alone would be ambiguous, e.g.
                 "VP Separate (Noise)" vs "Hearing Noise").

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

              GENERAL RULE: for a row with N non-zero defect columns, you MUST emit
              (N + 1) entries — one aggregate + N per-defect. NEVER emit only the
              aggregate and drop the per-defect rows. This is the #1 extraction
              failure across report types.

              Worked example (illustration of the rule — applies identically to ANY
              column set, not just the labels shown):

                Input row: Input=X, OK=Y, Total NG=Z, NG%=R%,
                           col_A = a (non-zero), col_B = b (non-zero), others = 0

                → Must emit EXACTLY 3 entries (aggregate + 2 per-defect):
                  [
                    { variable:V, inputQty:X, okQty:Y, ngTotal:Z, ngRate:R,
                      defectType:"",     defectCount:0 },
                    { variable:V, inputQty:X, okQty:Y, ngTotal:Z, ngRate:R,
                      defectType:col_A,  defectCount:a },
                    { variable:V, inputQty:X, okQty:Y, ngTotal:Z, ngRate:R,
                      defectType:col_B,  defectCount:b }
                  ]

                Concrete instance: a row "Normal, Input=195, NG=120,
                FRF=109, FRF+SPL=11, THD=0, No sound=0" yields 3 entries
                (aggregate, FRF=109, FRF+SPL=11). The FRF+SPL column is a compound
                label (rule 3a) — one entry, count=11, full label preserved.

                Emitting only the aggregate ({ngTotal:120, defectType:""}) and
                dropping the per-defect entries is WRONG regardless of which
                labels the report uses.

            Zero-count defect columns are ALWAYS omitted — never create zero-count entries.

            ══ STEP 6: SKIP AGGREGATE / DERIVED ROWS ══
            Skip rows of these kinds — they are NOT independent measurements:
            - Rows labeled "Total", "Grand Total", "Sum", sub-total rows.
            - PERCENT-ONLY continuation rows: identifiers (No/Date/Model/Type/
              Input/OK) are blank AND all numeric cells are X.X% values. These
              are derived percentages of the count row ABOVE them — do not emit
              a separate row. The count-row's ngRate already captures this.
            - Transcript rows prefixed with `(%)` from OCR output — always skip.
            - Transcript rows prefixed with `(cont)` — merge into the previous
              row, not a new measurement.

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

            **Dataset name mining (IMPORTANT):**
            The Dataset name "{{datasetName}}" itself often encodes signal about the
            report's intent (product code, process name, comparison type, lot info,
            date context). Parse it for meaningful tokens and fold them into the tags
            alongside what you read from the image(s). Ignore purely numeric or
            bookkeeping segments (sequence numbers, revision suffixes, raw timestamps).

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

            ══ STEP 8a: STRUCTURED CONTEXT EXTRACTION ══
            In addition to the free-form summary and key findings, extract FIVE
            focused context fields that power downstream Ask-AI queries. These are
            distinct from summary/keyFindings — keep each focused and short. If the
            report does not contain the info, leave the field as empty string "".

              • "purpose" — the explicit GOAL of the test. Draw from the report's
                "I. Purpose" / "Purpose" section if present; otherwise infer 1 short
                sentence. Examples: "Identify root cause of VP deform NG", "Validate
                new jig press VP/CD", "Evaluate lot change safety".

              • "testConditions" — WHAT was varied or changed in the test. Examples:
                "VP A2 moved to D3 Make Final", "Sub1cut 0.1mm new jig vs normal",
                "Pin offset -0.05 through +0.020", "Worker 1-7, before/after
                retraining". Include the comparison baseline when relevant.

              • "rootCause" — the identified CAUSE if the report concludes one, else "".
                Examples: "Day-shift D3-2B Audiobus machine process", "VP-separate
                check step introduces damage", "Coil offset #2/#4/#6 misalignment".

              • "decision" — the FINAL VERDICT from "IV. Decision" section if present,
                or the author's explicit conclusion. Examples: "Apply Frame clean by
                plasma to production", "New jig NOT suitable for production",
                "Further validation needed before rollout", "Exclude NG #2/#4/#6".

              • "recommendedAction" — concrete NEXT STEP if stated, else "".
                Examples: "Retrain day-shift operators", "Replace mold #6", "Increase
                UV drying to 15s total", "Switch to new grill design".

            Rules:
              - Each field: 1 short sentence max, English preferred.
              - Never fabricate. If a field is not in the report, emit "".
              - Do not repeat what's already in summary/keyFindings — these fields
                are FOCUSED facets, not a rewording.

            ══ STEP 9: SELF-CHECK BEFORE OUTPUT ══
            Before writing the final JSON, silently audit your measurement list
            against the source image. These checks are LABEL-AGNOSTIC — they apply
            to whatever defect columns the report happens to use.

              (a) DEFECT COVERAGE:
                  For every source-image row where Total NG > 0 AND the image shows
                  any non-zero value in per-defect columns, verify your output has
                  AT LEAST ONE entry with a non-empty defectType for that row
                  (matching variable, inputQty, ngTotal). If missing, ADD the
                  per-defect entries now. This is the single most common bug.
                  Applies to Normal/Baseline rows identically to Test rows.

              (b) COUNT SUM CHECK:
                  For each (variable, inputQty, ngTotal) group, sum the positive
                  defectCount values of entries where defectType is non-empty. The
                  sum should be ≥ ngTotal (one unit may have multiple defects so
                  the sum can exceed ngTotal, but it must not be smaller). If
                  smaller, a column was dropped — re-scan the row.

              (c) COMPOUND LABEL DISCIPLINE:
                  For every defect entry, check that the defectType string matches
                  EXACTLY one leaf header you enumerated in STEP 5(1). Any label
                  containing `+`, `&`, `/`, etc. is compound (one column) — its
                  value must not be split across two entries, and it must not be
                  conflated with an adjacent simple-labelled column.

              (d) LEAF COUNT SYMMETRY:
                  For any row, the number of distinct defectType values you emitted
                  for that row (including the empty aggregate) must be consistent
                  with the number of non-zero cells you read from the image. If
                  they differ, alignment is off.

            Fixing audit findings before output costs NOTHING; dropping defects
            is unrecoverable downstream.

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
              "tags": ["keyword1", "keyword2", "..."],
              "purpose": "",
              "testConditions": "",
              "rootCause": "",
              "decision": "",
              "recommendedAction": ""
            }
            """;

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        // 16384 tokens: reports with 30+ rows × multi-defect breakdown can exceed 8192
        // and get truncated, producing partial JSON that parses with missing per-defect rows.
        string raw = await CallWithContentAsync(blocks, 16384, ct, "claude-sonnet-4-6");
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
