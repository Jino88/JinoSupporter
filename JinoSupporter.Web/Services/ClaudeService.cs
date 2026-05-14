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
            "STEP 3: Handle merged cells and Data Input metadata.\n" +
            "- The TSV may already contain values expanded from real Excel merged ranges.\n" +
            "- Values may be prefixed with metadata like \"{bg=#FFFF00,merged=A1:A3}\" or \"〔bg=#FFFF00,merged=A1:A3〕\"; strip the metadata and use the following text as the cell value.\n" +
            "- For remaining empty cells, carry down only context columns such as Date, Model, Type, Line, No, section, or table group.\n" +
            "- Do NOT carry down numeric, OK, NG, total, rate, sample, or measurement columns.\n\n" +
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

        if (!string.IsNullOrWhiteSpace(rawExcelText))
        {
            blocks.Add(new JsonObject
            {
                ["type"] = "text",
                ["text"] =
                    "AUTHORITATIVE RAW EXCEL TEXT (tab-separated paste from the source " +
                    "workbook). When a cell value in the screenshot is ambiguous or OCR'd " +
                    "digits could be misread, PREFER the corresponding value from this raw " +
                    "text. Still transcribe every table structure from the image — use this " +
                    "only as a tiebreaker for cell values:\n```\n" + rawExcelText.Trim() + "\n```\n",
            });
        }

        string raw = await CallWithContentAsync(blocks, 64000, ct, "claude-sonnet-4-6");
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

            • ngRate: ALWAYS in PERCENT scale — number only, no "%" sign.
              "33.3%" → 33.3,  "100%" → 100.0,  "0.05%" → 0.05,  "5%" → 5.0
              NEVER store fractions. Do NOT output ngTotal/inputQty as a ratio —
              4/4 must be 100.0, not 1.0. If the report omits the % column but
              shows counts, compute (ngTotal / inputQty) * 100 and store THAT.

            • defectCategory mapping (use only these enums):
              assembly_defect / cosmetic_defect / function_spl / function_thd /
              function_hearing / wire_defect / magnetic_defect / rear_visual_damage / other.
              Header-group rule (resolves Audiobus ambiguity):
                - EVERY sub-column under a "NG AUDIOBUS" header  → function_spl
                  (SPL, SPL+RB, RB, No sound — ALL of them, including bare "RB")
                - EVERY sub-column under a "NG HEARING" header   → function_hearing
                  (Noise, Touch, Hearing)
                - "THD" column or "NG THD"                        → function_thd
                  (only literal THD — do NOT map RB / rear-buzz to function_thd)

            ══ REPORT-TYPE FIRST (v7) — CLASSIFY BEFORE WRITING ══

            ⚠️ Pick **reportType** FIRST — it changes which fields you fill.
            Pick exactly one:

            (a) `comparison_study`   — Normal vs Test (or Before vs After,
                old_lot vs new_lot). ONE main variant against ONE baseline.
                Default for "test new X" reports.
            (b) `multi_arm`          — 3+ variants compared (e.g. mold #4
                vs #6 vs #7 vs #8 / vendor A vs B vs C). When ≥3 distinct
                labels are compared on the same metric.
            (c) `doe_factorial`      — A 2-D factor grid (Temperature ×
                Tension, RPM × Voltage). Cells = (factor1 level, factor2
                level, OK/NG/value).
            (d) `reliability_validation` — Spec vs measured across multiple
                stress tests (high-temp, drop, salt-spray …). Verdict is
                `passed` / `failed`, not `improved` / `worsened`.
            (e) `trend_analysis`     — Time series. Weekly / monthly NG
                rate or worst-process trending over multiple periods.
            (f) `quality_log`        — Pure inspection record. No variant
                comparison, no spec gate. Just "we measured X on date Y".
                Verdict chip will be hidden in the UI.
            (g) `intervention_test`  — Single intervention measured but no
                obvious baseline pair (e.g. "tried bond X, NG rate = Y%").
                Falls back to 2-arm if a spec or prior value exists.

            ══ FIELD-BY-FIELD RULES (depend on reportType) ══

            Read this report as the matching study type. Output a STRUCTURED
            verdict + evidence rows + actions, plus type-specific payloads
            (`doeGrid` / `trendPoints` / multi-arm `comparisons[]`).

            Goal: ONE compact card the operator can read in five seconds.
            Verdict + headline at top; evidence numbers below; actions.
            ⚠️ Do NOT re-state the same facts in multiple fields. Each fact
               appears exactly once.

            🚫 NEVER prose. ✅ Numbers FIRST. Each text bullet ≤ 16 words.
               NO filler verbs ("결과", "확인", "수행", "통해", "되어").

            BASELINE SELECTION RULE (used for evidence rows):
            (1) If "Normal" exists in the same event, use Normal as baseline.
            (2) Else if "Before" / "Old" / "Standard" / "기존" exists, use that.
            (3) Else if A/B / Pos1/Pos2 → pairwise; set context.baselineReason
                to note low confidence.
            (4) DOE (Temperature × Tension × RPM combos) → no single baseline;
                emit one evidence row per axis or per best/worst cell, and
                explain in context.baselineReason.

            Field-by-field semantics:

            - verdict     — ENUM. Exactly one of (v7 — 7 values):
                            improved | worsened | partial | no_clear_effect |
                            inconclusive | passed | failed
              * improved          : variant beats baseline by a clear margin.
              * worsened          : variant is worse than baseline.
              * partial           : some metrics improved, others didn't.
              * no_clear_effect   : numbers indistinguishable (within ±0.5pp
                                    or sample N too small to call).
              * inconclusive      : data missing / DOE incomplete / aborted.
              * passed            : reliability_validation only — every test
                                    item within spec.
              * failed            : reliability_validation only — at least
                                    one item out of spec.
              For reportType=quality_log set verdict="" (empty) — no judgement.

            - headline    — Exactly one short sentence. Magnitude + direction.
                ✅ "VP press jig change → NG 8.3% → 2.7% (-5.6pp, improved)"
                ✅ "Plasma frame clean — function NG unchanged (3.0% both arms)"
                ❌ "전반적으로 양호한 결과를 보였습니다."

            - evidence    — Array of UP TO 4 rows. The KEY numeric comparisons.
                            Each row = one metric. Two paths:

              2-ARM path (default — reportType ∈ comparison_study /
              reliability_validation / intervention_test):
                * metric         — short label, e.g. "NG rate", "Hearing-Noise",
                                   "Gauss", "Tension".
                * baselineLabel  — "Normal", "Before", "Old lot", or specific
                                   condition. For reliability set "Spec".
                * baselineValue  — verbatim value w/ unit. "3.0% (3/100)",
                                   "477 G", "0.55 kgf", "> 1.530 kgf" (spec).
                * variantLabel   — "Test", "After", or specific variant.
                                   For reliability: "Measured".
                * variantValue   — same format as baselineValue.
                * deltaText      — display delta: "-5.6pp", "+0pp", "—".
                                   Use "—" when units aren't comparable or
                                   reportType=quality_log.
                * deltaSign      — ENUM: "up" | "down" | "no_change".
                                   ("down" = variant lower than baseline,
                                    regardless of whether lower is good.)
                * note           — optional, ≤8 words. e.g. "n=4 only — low conf".
                                   Leave empty if nothing to add.
                * comparisons    — leave empty / null in 2-arm path.

              MULTI-ARM path (reportType=multi_arm — 3+ variants on the same metric):
                * metric         — same as above.
                * comparisons    — array of ≥3, ≤8 entries. Order: baseline first
                                   if one exists, then variants by their natural
                                   order in the source (lot date, mold #, vendor).
                  · label      — "VP #4", "vendor A", "Reduce: 0.05".
                  · value      — verbatim value w/ unit. "10.4% (121/1160)".
                  · n          — sample size (integer).
                  · isBaseline — true on the baseline arm (Normal / spec / reference).
                  · isBest     — true on the best-performing arm.
                  · isWorst    — true on the worst-performing arm.
                * bestLabel / worstLabel — duplicate of best/worst.label for quick
                  UI access.
                * baselineLabel / baselineValue / variantLabel / variantValue / deltaText
                  — leave empty in multi-arm path. The UI ignores them.
                  ⚠ Exception: deltaText may hold "best vs worst" range string
                  like "+44.6pp range" for at-a-glance.

            - actions     — Array of UP TO 3 items, ordered by priority (1 = top).
                            Pure decisions/next-steps. NO restating numbers
                            (those are in evidence).
                * priority    — 1, 2, 3.
                * kind        — "action" (do this) | "investigate" (find out)
                                | "risk" (warn about).
                * text        — concrete imperative, ≤16 words.
                  ✅ "Apply X 81.65→81.60 to production line 6"
                  ✅ "Investigate root cause of yoke over-glue at sub-line"
                  ❌ "More testing needed."

            - doeGrid     — REQUIRED when reportType=doe_factorial, else null.
                * factor1Name   — "Temperature", "RPM", "Voltage" …
                * factor2Name   — second factor.
                * factor1Levels — ordered list of factor1 levels as STRINGS
                                  (e.g. ["380","390","400","410","420"]).
                * factor2Levels — same for factor2 (e.g. ["4","5","6","7","8"]).
                * cells         — one entry per (f1, f2) intersection.
                  · f1     — must match a factor1Levels entry verbatim.
                  · f2     — must match a factor2Levels entry verbatim.
                  · status — ENUM: "ok" | "ng" | "borderline" | "empty".
                  · value  — measured value or label (e.g. "7.611mm", "OK", "NG_melt").
                ⚠ Even when DOE is reportType, ALSO populate `evidence` with
                  AT MOST 2 summary rows (best cell, worst cell) so the
                  side-panel KPI still shows.

            - trendPoints — REQUIRED when reportType=trend_analysis, else null/empty.
                Array of points ordered chronologically.
                * label  — "Week 17", "2025-03", "Mar 2025", whatever the source uses.
                * value  — measured NG rate or metric for that period
                          (e.g. "8.3%", "120 ppm").
                * note   — optional 1-clause context ("freeze + holder change").
                Max 12 points; if more, summarise (skip intermediate).

            - context     — Optional object with 3 short string fields:
                * process        — name of the changed process/stage.
                  ✅ "Vision Bond glue inspection (BP/SM, MG/PT, Yoke)"
                * stage          — where the test ran.
                  ✅ "Sub Yoke 161016-D2 visual + E2-3A main-line function"
                * baselineReason — 1-line why this baseline was chosen.
                  ✅ "same-event Normal row present"
                  ✅ "no Normal; paired Type-1 vs Type-2 oversized samples"

            ══ TAGS ══
            Produce 4–10 English purpose-first tags (lowercase, hyphenated).
            Include product code, review type, main purpose, key comparison
            variable, intervention if present.

            ══ LEGACY NARRATIVE FIELDS ══
            ⚠️ The schema below still includes summary/keyFindings/purpose/
               testConditions/rootCause/decision/recommendedAction for
               backward compatibility. **Leave them as empty strings "".**
               All decision-relevant content goes into verdict/headline/
               evidence/actions/context above.

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
              "tags": ["keyword1", "keyword2", "..."],

              "reportType": "comparison_study | multi_arm | doe_factorial | reliability_validation | trend_analysis | quality_log | intervention_test",
              "verdict":    "improved | worsened | partial | no_clear_effect | inconclusive | passed | failed | \"\"",
              "headline":   "One short sentence with magnitude + direction.",
              "evidence": [
                {
                  "metric": "NG rate",
                  "baselineLabel": "Normal", "baselineValue": "3.0% (3/100)",
                  "variantLabel": "Type-2 Yoke", "variantValue": "3.0% (3/100)",
                  "deltaText": "+0pp", "deltaSign": "no_change", "note": "",
                  "comparisons": null,
                  "bestLabel": "", "worstLabel": ""
                }
                /* multi-arm row example:
                {
                  "metric": "VP bending rate",
                  "baselineLabel": "", "baselineValue": "",
                  "variantLabel": "", "variantValue": "",
                  "deltaText": "+58pp range", "deltaSign": "up", "note": "",
                  "comparisons": [
                    { "label":"VP #6", "value":"4.4% (53/1200)", "n":1200, "isBaseline":true,  "isBest":false, "isWorst":false },
                    { "label":"VP #7 improve", "value":"59.7% (689/1154)", "n":1154, "isBaseline":false, "isBest":false, "isWorst":true  },
                    { "label":"VP #9 improve", "value":"0.4% (5/1176)",   "n":1176, "isBaseline":false, "isBest":true,  "isWorst":false }
                  ],
                  "bestLabel":"VP #9 improve", "worstLabel":"VP #7 improve"
                } */
              ],
              "actions": [
                { "priority": 1, "kind": "action",      "text": "..." },
                { "priority": 2, "kind": "investigate", "text": "..." }
              ],
              "context": {
                "process": "",
                "stage": "",
                "baselineReason": ""
              },
              "doeGrid": null,
              /* doe example (only when reportType=doe_factorial):
              "doeGrid": {
                "factor1Name": "Temperature",
                "factor2Name": "Tension",
                "factor1Levels": ["380","390","400","410","420"],
                "factor2Levels": ["4","5","6","7","8"],
                "cells": [
                  { "f1":"390", "f2":"5", "status":"ok", "value":"7.611mm" },
                  { "f1":"380", "f2":"4", "status":"ng", "value":"NG_melt" }
                ]
              }, */
              "trendPoints": null,
              /* trend example (only when reportType=trend_analysis):
              "trendPoints": [
                { "label":"Week 17", "value":"8.3%",  "note":"hearing-noise dominant" },
                { "label":"Week 18", "value":"4.1%",  "note":"VP press change" },
                { "label":"Week 19", "value":"2.7%",  "note":"" }
              ], */

              "summary": "",
              "keyFindings": "",
              "purpose": "",
              "testConditions": "",
              "rootCause": "",
              "decision": "",
              "recommendedAction": ""
            }
            """;

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
            return new NormalizeResult { Summary = "JSON parse error — check Claude response (text-mode)." };
        }
    }

    // ── Normalize from images ─────────────────────────────────────────────────

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

        if (!string.IsNullOrWhiteSpace(rawText))
        {
            blocks.Add(new JsonObject
            {
                ["type"] = "text",
                ["text"] =
                    "AUTHORITATIVE RAW EXCEL TEXT (use this for exact cell values; " +
                    "prefer it over OCR'd numbers from the screenshot when they disagree):\n" +
                    "```\n" + rawText.Trim() + "\n```\n",
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
            - ngRate: FLOAT in PERCENT scale (number only, no "%" sign).
              Examples: "33.3%" → 33.3,  "100%" → 100.0,  "0.05%" → 0.05,  "5%" → 5.0
              NEVER store fractions. Do NOT compute ngTotal/inputQty as a ratio
              (e.g. 4/4 must NOT become 1.0 — the correct value is 100.0).
              If the report omits the rate column but shows counts, compute it as
              (ngTotal / inputQty) * 100 and store the PERCENT result.
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
            function_spl       → ANY sub-column under "NG AUDIOBUS" header
                                 (SPL, SPL+RB, RB, No sound — ALL of them,
                                 including bare "RB" and "SPL+RB")
            function_thd       → literal "THD" / "NG THD" column ONLY.
                                 RB (rear-buzz) under Audiobus is NOT THD.
            function_hearing   → ANY sub-column under "NG HEARING" header
                                 (Noise, Touch, Hearing)
            wire_defect        → wire offset, wire forming, wire cutting, wire clamp, wire pad offset, solder weak
            magnetic_defect    → Gauss low/NG
            rear_visual_damage → rear damage position N
            other              → anything else not clearly listed above

            ↳ Category decisions are driven by the PARENT MERGED-HEADER, not the
              bare sub-label. "RB" under "NG AUDIOBUS" is function_spl; "THD"
              under some other header is function_thd. Always look at the parent.

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
              - Each field: short scannable BULLETS ('\n'-separated), NOT prose.
                Numbers/values first; no narrative connectors. ≤12 words/bullet.
              - Use "key: from → to" or "key=value" form aggressively.
                Example testConditions: "X position: 81.65 → 81.60\nLot: 17/3, 2/4"
                Example rootCause: "VP lot 17/3 → 27.9% NG (vs 8.3% baseline)"
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
              "summary": "≤2 short bullets ('\\n'-separated). Numbers first. No prose.",
              "keyFindings": "5–8 short bullets ('\\n'-separated). 'metric: value' form, deltas with → / pp. ≤12 words each.",
              "tags": ["keyword1", "keyword2", "..."],
              "purpose": "",
              "testConditions": "",
              "rootCause": "",
              "decision": "",
              "recommendedAction": ""
            }
            """;

        blocks.Add(new JsonObject { ["type"] = "text", ["text"] = prompt });

        // 64000 tokens: large Normal-rich reports with 30+ rows × multi-defect breakdown
        // can exceed 32k and get truncated, producing partial JSON with missing per-defect rows.
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
                                              string answerLanguage = "English",
                                              CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(datasetsContext))
        {
            return new AskAiResult
            {
                Overall    = "No registered reports — cannot answer. Please save a report from Input Data first.",
                PerDataset = []
            };
        }

        string lang = string.IsNullOrWhiteSpace(answerLanguage) ? "English" : answerLanguage.Trim();

        string prompt = $$"""
            You are a manufacturing quality improvement assistant.

            A user has asked a question about a production problem. Answer it USING ONLY the information found in the registered dataset reports below.

            ══ STRICT RULES ══
            1. Do NOT use external/general knowledge. Only use facts present in the reports below.
            2. If no registered report contains relevant information, set "overall" to a short {{lang}} notice that no relevant data was found, and return an empty "perDataset" array. Do not invent an answer.
            3. Produce ONE entry in "perDataset" for EVERY dataset that genuinely contributes to the answer. Copy "datasetName" VERBATIM from the "Dataset:" header in the context (full string, including numeric prefixes and spaces).
            4. In each per-dataset "answer": explain in {{lang}} what this SPECIFIC dataset shows and how it addresses the user's question. Cite concrete values from that dataset only (NG rate, defect type, product type, date, specific findings). 2–5 sentences is ideal.
            5. Do NOT include datasets that are irrelevant to the question.
            5a. For NG-rate comparisons, judge improvement/worsening ONLY against the same-event Normal/Baseline/Control/Reference/Before/Old/OK row. Same-event means the same source sheet/table and same carried-forward Date/Model/Line/measurement type when those fields exist.
            5b. Merged Excel cells may appear blank in continuation rows. Treat blank Date/Model/Type cells below a visible value as carrying the visible value forward before pairing rows.
            5c. Use multiplicative relative change: (test_ng_rate / baseline_ng_rate - 1) * 100. Positive is worse; negative is improved. Do not use percentage-point subtraction as the verdict.
            5d. If no same-event baseline exists, do not say improved/worsened. Use ng_without_baseline style ranking, defect mix, source sheet, and sample size.
            5e. Respect report types: normal_comparison, ng_without_baseline, before_after_dimension, measurement_spec, defect_root_cause, lot_supplier_mold_comparison, process_condition_change, reliability_spec, doe_matrix, image_dependent, mixed. Answer in the matching shape: comparison table for comparison/process-change reports, spec/min/max/avg for measurement reports, cause/action/result for defect-cause reports.
            6. In "overall": give a 2–3 sentence {{lang}} synthesis across the per-dataset findings — top recommendations in priority order. If there is only one relevant dataset, you may leave "overall" empty.
            7. ALL human-readable text in the output ("overall" and every "answer") MUST be written in {{lang}}. Keep dataset names, product codes, defect type labels, and numeric values as-is.
            8. Return ONLY valid JSON — no markdown fences, no extra commentary.

            ══ OUTPUT JSON SCHEMA ══
            {
              "overall": "2–3 sentence {{lang}} overall recommendation across all datasets (may be empty).",
              "perDataset": [
                {
                  "datasetName": "<verbatim Dataset name>",
                  "answer": "{{lang}} dataset-specific answer with concrete numbers from this dataset."
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
                ?? new AskAiResult { Overall = "(Failed to parse response)" };
        }
        catch
        {
            return new AskAiResult { Overall = raw };
        }
    }

    // ── Translate analysis result (multi-field, one round-trip) ───────────────

    /// <summary>Translate the 7 narrative fields of a NormalizeResult into the
    /// target language in a single API call. Returns the original record's
    /// values for any field the model couldn't translate — never null fields.
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
            // legacy 7 fields — still translated for old-schema rows
            summary           = source.Summary           ?? "",
            keyFindings       = source.KeyFindings       ?? "",
            purpose           = source.Purpose           ?? "",
            testConditions    = source.TestConditions    ?? "",
            rootCause         = source.RootCause         ?? "",
            decision          = source.Decision          ?? "",
            recommendedAction = source.RecommendedAction ?? "",

            // v2 — new fields
            headline       = source.Headline ?? "",
            actionTexts    = actionTexts,
            contextProcess = source.Context?.Process        ?? "",
            contextStage   = source.Context?.Stage          ?? "",
            contextBaseline= source.Context?.BaselineReason ?? "",
        });

        string scriptRules = BuildOutputScriptRules(targetLanguage);

        string prompt = $$"""
            Translate every value in this JSON to {{targetLanguage}}. Keep keys
            unchanged. Preserve numbers, units, and product/part identifiers
            verbatim. If a field is empty, return it as an empty string.

            "actionTexts" is an array of strings — translate each element and
            return an array of the same length and order.

            {{scriptRules}}

            Output ONLY the translated JSON object — no markdown fence, no
            commentary, no extra keys.

            Input:
            {{inputJson}}
            """;

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

    // ── Translate ─────────────────────────────────────────────────────────────

    public async Task<string> TranslateAsync(string text,
                                             string targetLanguage,
                                             List<(string MediaType, string Base64)>? images = null,
                                             CancellationToken ct = default)
    {
        bool hasText  = !string.IsNullOrWhiteSpace(text);
        bool hasImage = images is { Count: > 0 };

        // Claude vision sometimes misreads ambiguous characters in screenshots and
        // emits them verbatim — e.g. dropping a Chinese hanzi into a Vietnamese line
        // ("Thêm hướng dẫn kh槽 V…"). Lock the output script with explicit rules so
        // every emitted character has to belong to the target language's writing
        // system. Translate semantically, never copy through unreadable glyphs.
        string scriptRules = BuildOutputScriptRules(targetLanguage);

        string instruction = (hasText, hasImage) switch
        {
            (true, false) =>
                $$"""
                Translate the following text to {{targetLanguage}}.

                {{scriptRules}}
                Return only the translation — no explanation, no original text, no commentary.

                {{text}}
                """,

            (false, true) =>
                $$"""
                Translate every visible text in the attached image(s) to {{targetLanguage}}.
                Preserve table structure, line breaks, and reading order.

                {{scriptRules}}
                Return only the translation — no explanation, no original text, no commentary.
                """,

            (true, true) =>
                $$"""
                Translate to {{targetLanguage}}. The user has provided BOTH source text and
                reference image(s); use the images as context for ambiguous terms but treat
                the text below as the primary content to translate.

                {{scriptRules}}
                Return only the translation — no explanation, no original text, no commentary.

                {{text}}
                """,

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
    /// misreads — e.g. dropping a Chinese hanzi into a Vietnamese line because
    /// the source character was ambiguous in the screenshot.</summary>
    private static string BuildOutputScriptRules(string targetLanguage)
    {
        string scriptLine = targetLanguage switch
        {
            "Korean"                => "Korean uses Hangul (한글) plus standard Latin punctuation/numerals. Do NOT include Chinese hanzi, Japanese kana, or any other non-Hangul script.",
            "English"               => "English uses the unaccented Latin alphabet (A-Z, a-z) plus standard punctuation/numerals. Do NOT include Chinese, Korean, Japanese, or any non-Latin characters.",
            "Japanese"              => "Japanese uses Hiragana, Katakana, and Kanji (CJK characters that are part of the standard Japanese writing system). Do NOT include Hangul or non-Japanese scripts.",
            "Chinese (Simplified)"  => "Chinese (Simplified) uses Simplified Han characters plus standard punctuation/numerals. Do NOT include Traditional-only characters, Hangul, kana, or other non-Chinese scripts.",
            "Chinese (Traditional)" => "Chinese (Traditional) uses Traditional Han characters plus standard punctuation/numerals. Do NOT include Simplified-only characters, Hangul, kana, or other non-Chinese scripts.",
            "Vietnamese"            => "Vietnamese uses the Latin alphabet with diacritics (à á ả ã ạ ă â ê ô ơ ư đ etc.) plus standard punctuation/numerals. Do NOT include Chinese hanzi, Korean Hangul, Japanese kana, or any non-Latin character — every glyph must be a valid Vietnamese letter, digit, or punctuation mark.",
            "Spanish"               => "Spanish uses the Latin alphabet with diacritics (á é í ó ú ñ ü) plus ¿ ¡ and standard punctuation/numerals. Do NOT include any non-Latin character.",
            "French"                => "French uses the Latin alphabet with diacritics (à â ç é è ê ë î ï ô ù û ü ÿ œ æ) plus standard punctuation/numerals. Do NOT include any non-Latin character.",
            "German"                => "German uses the Latin alphabet plus ä ö ü ß and standard punctuation/numerals. Do NOT include any non-Latin character.",
            _                       => $"Output must use only the standard writing system for {targetLanguage}. Do NOT mix in characters from other scripts (Chinese hanzi, Korean Hangul, Japanese kana, etc.).",
        };

        return $$"""
            Output script rules:
            - {{scriptLine}}
            - If a character in the source is unreadable or ambiguous, do not copy it verbatim. Translate the surrounding context semantically and use your best linguistic judgment instead.
            - Proper nouns / brand names / product codes that are already in the Latin alphabet may stay as-is when the target uses Latin script; otherwise transliterate them into the target script.
            """;
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
