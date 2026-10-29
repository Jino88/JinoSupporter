using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Review report writer. The user drops in notes, pasted Excel cells and pictures; the Codex API
/// writes a report whose sections follow the WorkHub review-data normalization fields (purpose,
/// changed factors, control/test conditions, result tables with units, source conclusion, tags)
/// and places each picture under the section it belongs to. The report is kept in its own SQLite
/// file and exported as .xlsx, the format WorkHub registers and normalizes.
/// </summary>
public sealed class ReviewReportService(CodexApiService codex, ILogger<ReviewReportService> log)
{
    public const string Model = CodexApiService.LunaModel;
    // Copying numbers into a fixed JSON shape needs little reasoning; low keeps the call short.
    public const string Effort = "low";

    private const int MaxImages = 12;
    private const long ImageBudgetBytes = 1_500_000;

    public static readonly string[] Languages = ["Same as input", "English", "Korean", "Vietnamese"];
    public static readonly string[] Sections = ["purpose", "conditions", "results", "conclusion", "appendix"];

    private static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        WriteIndented = false,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        Converters = { new LenientStringConverter(), new MetricValuesConverter() },
    };

    /// <summary>The AI sometimes answers a table cell as a number or null; read those as text.</summary>
    private sealed class LenientStringConverter : JsonConverter<string>
    {
        public override bool HandleNull => true;

        public override string Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options) =>
            reader.TokenType switch
            {
                JsonTokenType.String => reader.GetString() ?? "",
                JsonTokenType.Number => Encoding.UTF8.GetString(reader.ValueSpan),
                JsonTokenType.True => "true",
                JsonTokenType.False => "false",
                JsonTokenType.Null => "",
                _ => JsonDocument.ParseValue(ref reader).RootElement.GetRawText(),
            };

        public override void Write(Utf8JsonWriter writer, string value, JsonSerializerOptions options) =>
            writer.WriteStringValue(value ?? "");
    }

    /// <summary>
    /// Metric values: the AI answers the short form { "C1": "38" } (fewer output tokens, faster);
    /// saved reports keep the list form [ { "arm_id": "C1", "value": "38" } ]. Both are read.
    /// </summary>
    private sealed class MetricValuesConverter : JsonConverter<List<ReviewReportValue>>
    {
        public override List<ReviewReportValue> Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            using JsonDocument d = JsonDocument.ParseValue(ref reader);
            var list = new List<ReviewReportValue>();
            static string Text(JsonElement e) => e.ValueKind switch
            {
                JsonValueKind.String => e.GetString() ?? "",
                JsonValueKind.Null or JsonValueKind.Undefined => "",
                _ => e.GetRawText(),
            };
            if (d.RootElement.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty p in d.RootElement.EnumerateObject())
                    list.Add(new ReviewReportValue { ArmId = p.Name, Value = Text(p.Value) });
            }
            else if (d.RootElement.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement e in d.RootElement.EnumerateArray())
                {
                    if (e.ValueKind != JsonValueKind.Object) continue;
                    string arm = "", value = "";
                    foreach (JsonProperty p in e.EnumerateObject())
                    {
                        if (p.Name.Equals("arm_id", StringComparison.OrdinalIgnoreCase)) arm = Text(p.Value);
                        else if (p.Name.Equals("value", StringComparison.OrdinalIgnoreCase)) value = Text(p.Value);
                    }
                    list.Add(new ReviewReportValue { ArmId = arm, Value = value });
                }
            }
            return list;
        }

        public override void Write(Utf8JsonWriter writer, List<ReviewReportValue> value, JsonSerializerOptions options)
        {
            writer.WriteStartArray();
            foreach (ReviewReportValue v in value)
            {
                writer.WriteStartObject();
                writer.WriteString("arm_id", v.ArmId);
                writer.WriteString("value", v.Value);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }
    }

    private static readonly object SchemaGate = new();
    private static bool _schemaReady;

    public bool IsConfigured => codex.IsConfigured;

    public static string DbPath => AppStoragePaths.Combine("review-report.db");

    // ── AI ──────────────────────────────────────────────────────────────────

    public async Task<ReviewReportDoc> GenerateAsync(
        string inputText,
        IReadOnlyList<ReviewReportImage> images,
        string language,
        string extraInstruction,
        CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(inputText) && images.Count == 0)
            throw new InvalidOperationException("Enter the content or add a picture first.");

        var content = new JsonArray();
        var head = new StringBuilder();
        head.AppendLine($"Report language: {language}");
        if (!string.IsNullOrWhiteSpace(extraInstruction))
            head.AppendLine($"Extra instruction from the author: {extraInstruction.Trim()}");
        head.AppendLine($"Pictures attached: {Math.Min(images.Count, MaxImages)} (numbered IMG 1..{Math.Min(images.Count, MaxImages)} in the order below)");
        head.AppendLine();
        head.AppendLine("=== CONTENT (notes, pasted Excel cells as tab-separated text) ===");
        head.AppendLine(string.IsNullOrWhiteSpace(inputText) ? "(none)" : inputText.Trim());
        content.Add(new JsonObject { ["type"] = "input_text", ["text"] = head.ToString() });

        return await CallAsync(Instructions, content, images, ct);
    }

    /// <summary>
    /// Continues an existing report: the AI gets the current report, the original content and the
    /// earlier requests, applies only the new change request and answers the whole report again.
    /// </summary>
    public async Task<ReviewReportDoc> ReviseAsync(ReviewReport report, string request, CancellationToken ct = default)
    {
        if (report.Doc is null) throw new InvalidOperationException("Write the report first.");
        if (string.IsNullOrWhiteSpace(request)) throw new InvalidOperationException("Enter what to change.");

        var head = new StringBuilder();
        head.AppendLine($"Report language: {report.Language}");
        head.AppendLine($"Pictures attached: {Math.Min(report.Images.Count, MaxImages)} (numbered IMG 1..{Math.Min(report.Images.Count, MaxImages)} in the order below)");
        head.AppendLine();
        head.AppendLine("=== CURRENT REPORT (JSON) ===");
        head.AppendLine(JsonSerializer.Serialize(report.Doc, Json));
        head.AppendLine();
        head.AppendLine("=== ORIGINAL CONTENT (notes, pasted Excel cells) ===");
        head.AppendLine(string.IsNullOrWhiteSpace(report.InputText) ? "(none)" : report.InputText.Trim());
        if (report.Revisions.Count > 0)
        {
            head.AppendLine();
            head.AppendLine("=== EARLIER CHANGE REQUESTS (already applied) ===");
            foreach (ReviewReportRevision r in report.Revisions) head.AppendLine("- " + r.Request.Trim());
        }
        head.AppendLine();
        head.AppendLine("=== CHANGE REQUEST ===");
        head.AppendLine(request.Trim());
        var content = new JsonArray { new JsonObject { ["type"] = "input_text", ["text"] = head.ToString() } };
        return await CallAsync(Instructions + RevisionInstructions, content, report.Images, ct);
    }

    private const string RevisionInstructions = """

        REVISION: you are given the current report as JSON and a change request from the author.
        Apply the change request and answer the WHOLE report again as one JSON object with the same
        keys. Change only what the request asks for (and what must change with it); keep every other
        field, number, condition id, metric id and picture placement exactly as it is. The rules above
        still hold: numbers come from the original content or pictures, never invented. Earlier
        change requests are already applied; do not undo them.
        """;

    private async Task<ReviewReportDoc> CallAsync(string instructions, JsonArray content, IReadOnlyList<ReviewReportImage> images, CancellationToken ct)
    {
        foreach (ReviewReportImage img in images.Take(MaxImages))
        {
            string b64 = Convert.ToBase64String(img.Data);
            string mediaType = img.MediaType;
            if (OperatingSystem.IsWindowsVersionAtLeast(6, 1) && img.Data.LongLength > ImageBudgetBytes)
            {
                try { (b64, mediaType) = ImageCompressor.CompressIfLarge(b64, mediaType, ImageBudgetBytes); }
                catch (Exception ex) { log.LogWarning(ex, "Could not shrink review report image {No}", img.No); }
            }
            content.Add(new JsonObject { ["type"] = "input_text", ["text"] = $"IMG {img.No} ({img.FileName})" });
            content.Add(new JsonObject { ["type"] = "input_image", ["image_url"] = $"data:{mediaType};base64,{b64}" });
        }

        var sw = System.Diagnostics.Stopwatch.StartNew();
        CodexTextResult result = await codex.GenerateAsync(instructions, content, Model, Effort, 16000, ct);
        codex.RecordUsage(result.Usage);
        log.LogInformation("Review report generated in {Seconds:0}s: chars={Chars} images={Images} in={In} out={Out}",
            sw.Elapsed.TotalSeconds, result.Text.Length, images.Count, result.Usage?.InputTokens, result.Usage?.OutputTokens);

        ReviewReportDoc doc = ParseDoc(result.Text);
        // Keep only placements that point at a real picture; any picture the AI left out goes to the appendix.
        var known = images.Select(i => i.No).ToHashSet();
        doc.Images = doc.Images.Where(p => known.Contains(p.No)).GroupBy(p => p.No).Select(g => g.First()).ToList();
        foreach (ReviewReportImage img in images.Where(i => doc.Images.All(p => p.No != i.No)))
            doc.Images.Add(new ReviewReportImagePlacement { No = img.No, Caption = img.FileName, Section = "appendix" });
        foreach (var p in doc.Images.Where(p => !Sections.Contains(p.Section)))
            p.Section = "appendix";
        return doc;
    }

    private const string Instructions = """
        You write a manufacturing review report (test / experiment / defect investigation) from the
        author's notes, pasted Excel cells and pictures. The report is later registered in a review
        database where an AI normalizes it into fixed fields so related reviews can be searched and
        compared. Write it so that normalization is easy and lossless:

        - Every section maps to one normalization field: purpose, changed factors (baseline -> changed),
          conditions/arms (control vs test), metrics (one value per condition), the author's own
          conclusion, limitations, tags. Each fact appears in ONE place only.
        - Copy numbers, units, model names, part numbers and dates EXACTLY as they appear in the content
          or pictures. Never compute, round, estimate or invent a value. Do not add your own
          recommendation; the conclusion is the author's conclusion (or, if the author gave none, a
          plain statement of what the numbers show, e.g. "New Jig total NG rate 0.8% vs Old Jig 5.7%").

        Pasted Excel cells lose their layout. Rebuild it before writing:
        - A merged cell arrives as a value on the first row and blanks below it. When a column such as
          date, lot, model or group has a value on one row and a blank on the next row(s) of the same
          block, it was merged: apply the value to every row of the block.
        - A cell that holds two stacked values (count on top, rate below) arrives as an extra line
          under the row (e.g. "5.7%  0.0%" under the counts 38 and 0). Attach those values to the row
          above, column by column, as the rate of each count.

        Conditions (arms): one per compared item (e.g. Old Jig = control, New Jig = test). Fill id, label,
        role (control | test | other), date, model, table and details. "table" names the review the
        condition belongs to (e.g. "Laser cutting 6.72 -> 6.70"); conditions compared with each other
        share one table name, a separate review or test gets another. "details" describes only the setup
        (what is different about this arm). NEVER put result numbers in conditions.

        Results are stored normalized, the way the review database keeps them: a list of metrics,
        each with one value per condition (arm). Do NOT build a table yourself.
        - One metric per measured item, in the source order: Input, OK, each defect type, total NG,
          total NG rate, measurements (mm, mW/cm², s ...).
        - id: short snake_case ("input", "ng_vp_offset", "ng_vp_offset_rate"); name: as the author wrote it;
          unit: pcs, %, mm ... ("" when none); kind: "input" (input quantity), "ok", "ng" (the count of
          one defect type), "total_ng", "rate" or "value" (a measurement).
        - Do not compute rates or ppm: the app shows each NG count with its NG rate in ppm.
        - A rate stacked under a count in the source (count on top, rate below in the same cell area)
          is its own metric with kind "rate" and rate_of = the id of that count metric.
        - values: an object from condition id to value, e.g. { "C1": "38", "C2": "0" }. Copy the value
          exactly as displayed ("5.7%", "0.12", "1,200") but only the number: no unit (it goes in
          "unit"), no "≈", no words. Skip conditions without a value.

        Report layout (the author's Excel report form): I. Purpose, II. Content (what was done, the
        setup pictures), III. Result (one result table per comparison + defect pictures), IV. Conclusion.
        - content: short "- ..." lines saying what was done (e.g. "Test New JIG and check NG rate").
        - result_title: the heading above the result table when there is one review (e.g. "Result
          check New Jig"); with several reviews each table is headed by its condition "table" name.
        - The result tables are drawn from the metrics: one table per review, one row per condition,
          one column per metric. Conditions of one review measured on different items get separate
          tables, so give each condition only the metrics really measured for it.

        Pictures: read every picture. When a picture is a table, its numbers go into the metrics and
        the picture is still placed. Each picture gets:
        - section: "purpose" (background, the problem), "conditions" (setup, jig, sample, process),
          "results" (defect photos, charts, data after the test), "conclusion", or "appendix".
        - group: pictures that belong side by side share the same short group name (e.g. "Jig
          comparison", "NG VP cutting offset", "No cutting"), in the order they should appear.
        - caption: a short label for that one picture (e.g. "Old Jig", "New Jig", "Offset sample 1").

        Tags: model codes, process names, defect types, jig/test names (short, 3-10 items).
        Write all prose in the requested report language ("Same as input" = the main language of the
        content). Keep codes, units and the author's terms as written.

        Answer with ONE JSON object only, no markdown fence, with exactly these keys:
        {
          "title": "",
          "product_model": "",
          "report_date": "",
          "purpose": "",
          "background": "",
          "content": [ "" ],
          "result_title": "",
          "changed_factors": [ { "name": "", "baseline": "", "changed": "" } ],
          "conditions": [ { "id": "C1", "label": "", "role": "control|test|other", "date": "", "model": "", "table": "", "details": "" } ],
          "metrics": [ { "id": "", "name": "", "unit": "", "kind": "input|ok|ng|total_ng|rate|value", "rate_of": "", "values": { "C1": "", "C2": "" } } ],
          "findings": [ "" ],
          "conclusion": "",
          "limitations": [ "" ],
          "next_actions": [ "" ],
          "tags": [ "" ],
          "images": [ { "no": 1, "caption": "", "group": "", "section": "purpose|conditions|results|conclusion|appendix" } ]
        }
        Use empty strings / empty arrays for anything the content does not give.
        """;

    internal static ReviewReportDoc ParseDoc(string text)
    {
        string s = text.Trim();
        int first = s.IndexOf('{');
        int last = s.LastIndexOf('}');
        if (first < 0 || last <= first)
            throw new InvalidOperationException("The AI answer was not a JSON report: " + (s.Length > 300 ? s[..300] + "…" : s));
        try
        {
            ReviewReportDoc? doc = JsonSerializer.Deserialize<ReviewReportDoc>(s[first..(last + 1)], Json);
            return doc ?? throw new InvalidOperationException("The AI answer was empty.");
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException("Could not read the AI report JSON: " + ex.Message);
        }
    }

    // ── Storage ─────────────────────────────────────────────────────────────

    private static SqliteConnection Open()
    {
        string path = DbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var conn = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = path,
            Mode = SqliteOpenMode.ReadWriteCreate,
            DefaultTimeout = 30,
        }.ToString());
        conn.Open();
        lock (SchemaGate)
        {
            if (!_schemaReady)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    PRAGMA journal_mode = WAL;
                    CREATE TABLE IF NOT EXISTS reports (
                        id TEXT PRIMARY KEY,
                        title TEXT NOT NULL DEFAULT '',
                        created_by TEXT NOT NULL DEFAULT '',
                        created_at TEXT NOT NULL,
                        updated_at TEXT NOT NULL,
                        language TEXT NOT NULL DEFAULT '',
                        input_text TEXT NOT NULL DEFAULT '',
                        extra_instruction TEXT NOT NULL DEFAULT '',
                        report_json TEXT,
                        model TEXT NOT NULL DEFAULT ''
                    );
                    CREATE TABLE IF NOT EXISTS report_images (
                        report_id TEXT NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                        no INTEGER NOT NULL,
                        file_name TEXT NOT NULL DEFAULT '',
                        media_type TEXT NOT NULL DEFAULT 'image/png',
                        data BLOB NOT NULL,
                        PRIMARY KEY (report_id, no)
                    );
                    -- Result values, one row per (metric, condition), rewritten on every save.
                    CREATE TABLE IF NOT EXISTS report_values (
                        report_id TEXT NOT NULL REFERENCES reports(id) ON DELETE CASCADE,
                        metric_id TEXT NOT NULL,
                        arm_id TEXT NOT NULL,
                        metric_name TEXT NOT NULL DEFAULT '',
                        unit TEXT NOT NULL DEFAULT '',
                        kind TEXT NOT NULL DEFAULT '',
                        rate_of TEXT NOT NULL DEFAULT '',
                        arm_label TEXT NOT NULL DEFAULT '',
                        arm_role TEXT NOT NULL DEFAULT '',
                        source_display TEXT NOT NULL DEFAULT '',
                        number REAL,
                        PRIMARY KEY (report_id, metric_id, arm_id)
                    );
                    """;
                cmd.ExecuteNonQuery();
                using (var col = conn.CreateCommand())
                {
                    col.CommandText = "SELECT COUNT(*) FROM pragma_table_info('reports') WHERE name = 'revisions_json'";
                    if (Convert.ToInt64(col.ExecuteScalar()) == 0)
                    {
                        col.CommandText = "ALTER TABLE reports ADD COLUMN revisions_json TEXT NOT NULL DEFAULT '[]'";
                        col.ExecuteNonQuery();
                    }
                }
                _schemaReady = true;
            }
        }
        using (var fk = conn.CreateCommand())
        {
            fk.CommandText = "PRAGMA foreign_keys = ON;";
            fk.ExecuteNonQuery();
        }
        return conn;
    }

    public List<ReviewReportSummary> List()
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT id, title, created_by, updated_at FROM reports ORDER BY updated_at DESC";
        var list = new List<ReviewReportSummary>();
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new ReviewReportSummary(r.GetString(0), r.GetString(1), r.GetString(2), ParseTime(r.GetString(3))));
        return list;
    }

    public ReviewReport? Get(string id)
    {
        using var conn = Open();
        ReviewReport report;
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT id, title, created_by, created_at, updated_at, language, input_text, extra_instruction, report_json, model, revisions_json
                FROM reports WHERE id = @id
                """;
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            if (!r.Read()) return null;
            report = new ReviewReport
            {
                Id = r.GetString(0),
                Title = r.GetString(1),
                CreatedBy = r.GetString(2),
                CreatedAt = ParseTime(r.GetString(3)),
                UpdatedAt = ParseTime(r.GetString(4)),
                Language = r.GetString(5),
                InputText = r.GetString(6),
                ExtraInstruction = r.GetString(7),
                Doc = r.IsDBNull(8) ? null : JsonSerializer.Deserialize<ReviewReportDoc>(r.GetString(8), Json),
                Model = r.GetString(9),
                Revisions = JsonSerializer.Deserialize<List<ReviewReportRevision>>(r.IsDBNull(10) ? "[]" : r.GetString(10), Json) ?? [],
            };
        }
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT no, file_name, media_type, data FROM report_images WHERE report_id = @id ORDER BY no";
            cmd.Parameters.AddWithValue("@id", id);
            using var r = cmd.ExecuteReader();
            while (r.Read())
                report.Images.Add(new ReviewReportImage(r.GetInt32(0), r.GetString(1), r.GetString(2), (byte[])r["data"]));
        }
        return report;
    }

    public void Save(ReviewReport report)
    {
        if (string.IsNullOrEmpty(report.Id)) report.Id = Guid.NewGuid().ToString("N");
        report.UpdatedAt = DateTime.Now;
        if (report.CreatedAt == default) report.CreatedAt = report.UpdatedAt;
        report.Title = FirstNonBlank(report.Doc?.Title, report.Title, "Untitled report");

        using var conn = Open();
        using var tx = conn.BeginTransaction();
        using (var cmd = conn.CreateCommand())
        {
            cmd.Transaction = tx;
            cmd.CommandText = """
                INSERT INTO reports (id, title, created_by, created_at, updated_at, language, input_text, extra_instruction, report_json, model, revisions_json)
                VALUES (@id, @title, @by, @created, @updated, @lang, @input, @extra, @json, @model, @revisions)
                ON CONFLICT(id) DO UPDATE SET
                    title = excluded.title, updated_at = excluded.updated_at, language = excluded.language,
                    input_text = excluded.input_text, extra_instruction = excluded.extra_instruction,
                    report_json = excluded.report_json, model = excluded.model, revisions_json = excluded.revisions_json
                """;
            cmd.Parameters.AddWithValue("@id", report.Id);
            cmd.Parameters.AddWithValue("@title", report.Title);
            cmd.Parameters.AddWithValue("@by", report.CreatedBy);
            cmd.Parameters.AddWithValue("@created", FormatTime(report.CreatedAt));
            cmd.Parameters.AddWithValue("@updated", FormatTime(report.UpdatedAt));
            cmd.Parameters.AddWithValue("@lang", report.Language);
            cmd.Parameters.AddWithValue("@input", report.InputText);
            cmd.Parameters.AddWithValue("@extra", report.ExtraInstruction);
            cmd.Parameters.AddWithValue("@json", report.Doc is null ? DBNull.Value : JsonSerializer.Serialize(report.Doc, Json));
            cmd.Parameters.AddWithValue("@model", report.Model);
            cmd.Parameters.AddWithValue("@revisions", JsonSerializer.Serialize(report.Revisions, Json));
            cmd.ExecuteNonQuery();
        }
        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM report_images WHERE report_id = @id";
            del.Parameters.AddWithValue("@id", report.Id);
            del.ExecuteNonQuery();
        }
        foreach (ReviewReportImage img in report.Images)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO report_images (report_id, no, file_name, media_type, data) VALUES (@id, @no, @name, @mt, @data)";
            ins.Parameters.AddWithValue("@id", report.Id);
            ins.Parameters.AddWithValue("@no", img.No);
            ins.Parameters.AddWithValue("@name", img.FileName);
            ins.Parameters.AddWithValue("@mt", img.MediaType);
            ins.Parameters.Add("@data", SqliteType.Blob).Value = img.Data;
            ins.ExecuteNonQuery();
        }
        using (var delValues = conn.CreateCommand())
        {
            delValues.Transaction = tx;
            delValues.CommandText = "DELETE FROM report_values WHERE report_id = @id";
            delValues.Parameters.AddWithValue("@id", report.Id);
            delValues.ExecuteNonQuery();
        }
        if (report.Doc is { } doc)
        {
            var arms = ResultArms(doc).ToDictionary(a => a.Id.Trim(), StringComparer.OrdinalIgnoreCase);
            var written = new HashSet<(string, string)>();
            foreach (ReviewReportMetric m in doc.Metrics.Where(m => !string.IsNullOrWhiteSpace(m.Id)))
            {
                foreach (ReviewReportValue v in m.Values)
                {
                    string armId = v.ArmId.Trim();
                    if (armId.Length == 0 || string.IsNullOrWhiteSpace(v.Value) || !written.Add((m.Id.Trim(), armId))) continue;
                    ReviewReportCondition? arm = arms.GetValueOrDefault(armId);
                    using var ins = conn.CreateCommand();
                    ins.Transaction = tx;
                    ins.CommandText = """
                        INSERT INTO report_values (report_id, metric_id, arm_id, metric_name, unit, kind, rate_of, arm_label, arm_role, source_display, number)
                        VALUES (@id, @m, @a, @name, @unit, @kind, @of, @label, @role, @shown, @num)
                        """;
                    ins.Parameters.AddWithValue("@id", report.Id);
                    ins.Parameters.AddWithValue("@m", m.Id.Trim());
                    ins.Parameters.AddWithValue("@a", armId);
                    ins.Parameters.AddWithValue("@name", m.Name);
                    ins.Parameters.AddWithValue("@unit", m.Unit);
                    ins.Parameters.AddWithValue("@kind", m.Kind);
                    ins.Parameters.AddWithValue("@of", m.RateOf);
                    ins.Parameters.AddWithValue("@label", arm?.Label ?? "");
                    ins.Parameters.AddWithValue("@role", arm?.Role ?? "");
                    ins.Parameters.AddWithValue("@shown", v.Value.Trim());
                    ins.Parameters.AddWithValue("@num", ParseNumber(v.Value) is double d ? d : DBNull.Value);
                    ins.ExecuteNonQuery();
                }
            }
        }
        tx.Commit();
    }

    /// <summary>"1,200" → 1200, "5.7%" → 0.057 (a rate as a fraction, like the review database); otherwise null.</summary>
    /// <remarks>
    /// Values come as written on the sheet or the note: "111 pcs", "≈ 5,4%", "1,200", "1.234,5".
    /// The first number is read; a comma is a thousands separator only in groups of three digits,
    /// otherwise it is the decimal mark.
    /// </remarks>
    internal static double? ParseNumber(string? value)
    {
        string v = (value ?? "").Trim();
        var m = System.Text.RegularExpressions.Regex.Match(v, @"[-+]?\d[\d.,]*");
        if (!m.Success) return null;
        string n = m.Value.TrimEnd('.', ',');
        bool percent = v[(m.Index + m.Length)..].TrimStart().StartsWith('%');
        if (n.Contains(',') && n.Contains('.'))
            n = n.LastIndexOf(',') > n.LastIndexOf('.') ? n.Replace(".", "").Replace(',', '.') : n.Replace(",", "");
        else if (n.Contains(','))
            n = System.Text.RegularExpressions.Regex.IsMatch(n, @"^[-+]?\d{1,3}(,\d{3})+$") ? n.Replace(",", "") : n.Replace(',', '.');
        if (!double.TryParse(n, NumberStyles.Float, CultureInfo.InvariantCulture, out double d)) return null;
        return percent ? d / 100.0 : d;
    }

    public void Delete(string id)
    {
        using var conn = Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM report_values WHERE report_id = @id; DELETE FROM report_images WHERE report_id = @id; DELETE FROM reports WHERE id = @id;";
        cmd.Parameters.AddWithValue("@id", id);
        cmd.ExecuteNonQuery();
    }

    private static string FormatTime(DateTime t) => t.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
    private static DateTime ParseTime(string s) =>
        DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var t) ? t : default;

    private static string FirstNonBlank(params string?[] values) =>
        values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? "";

    // ── Excel export (the file WorkHub registers) ───────────────────────────

    // The author's Excel report form: grey "I. Purpose" bars, the result table in light blue with
    // the rate in red under its count, the test row in yellow, a yellow label bar under pictures.
    private const int LastCol = 13;
    private const double ColPx = 82.0;      // a column of width 11
    private const double PxPerRow = 20.0;
    private const string FormFont = "Malgun Gothic";
    private static readonly XLColor SectionGrey = XLColor.FromHtml("#BFBFBF");
    private static readonly XLColor HeaderBlue = XLColor.FromHtml("#C5D9F1");
    private static readonly XLColor TestYellow = XLColor.FromHtml("#FFFF00");
    private static readonly XLColor RateRed = XLColor.FromHtml("#FF0000");

    public static byte[] ExportXlsx(ReviewReport report)
    {
        ReviewReportDoc doc = report.Doc ?? new ReviewReportDoc();
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Report");
        ws.ShowGridLines = false;
        ws.Column(1).Width = 2;
        for (int c = 2; c <= LastCol; c++) ws.Column(c).Width = 11;
        ws.Style.Font.FontName = FormFont;
        ws.Style.Font.FontSize = 11;
        ws.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        int row = 1;
        var titleRange = ws.Range(row, 1, row, LastCol).Merge();
        titleRange.Value = FirstNonBlank(doc.Title, report.Title);
        titleRange.Style.Font.Bold = true;
        titleRange.Style.Font.FontSize = 16;
        titleRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Row(row).Height = 30;
        row++;
        string info = string.Join("    ", new[]
        {
            Label("Model", doc.ProductModel), Label("Date", doc.ReportDate), Label("Author", report.CreatedBy),
        }.Where(x => x.Length > 0));
        if (info.Length > 0)
        {
            var infoRange = ws.Range(row, 1, row, LastCol).Merge();
            infoRange.Value = info;
            infoRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            infoRange.Style.Font.FontSize = 10;
            row++;
        }
        row++;

        var picsBySection = doc.Images.GroupBy(p => p.Section).ToDictionary(g => g.Key, g => g.ToList());
        var imageByNo = report.Images.ToDictionary(i => i.No);

        row = Section(ws, row, "I. Purpose.");
        row = Paragraph(ws, row, Dash(doc.Purpose));
        row = Paragraph(ws, row, Dash(doc.Background));
        row = Pictures(ws, row + 1, picsBySection.GetValueOrDefault("purpose"), imageByNo);
        row++;

        row = Section(ws, row, "II. Content.");
        foreach (string line in doc.Content.Where(l => !string.IsNullOrWhiteSpace(l)))
            row = Paragraph(ws, row, Dash(line));
        foreach (ReviewReportFactor f in doc.ChangedFactors.Where(f => !string.IsNullOrWhiteSpace(f.Name)))
            row = Paragraph(ws, row, $"- {f.Name}: {f.Baseline} → {f.Changed}");
        foreach (ReviewReportCondition c in doc.Conditions.Where(c => !string.IsNullOrWhiteSpace(c.Details)))
            row = Paragraph(ws, row, $"- {c.Label}: {c.Details}");
        row = Pictures(ws, row + 1, picsBySection.GetValueOrDefault("conditions"), imageByNo);
        row++;

        row = Section(ws, row, "III. Result");
        int no = 1;
        foreach (ReviewReportResultTable table in ResultTables(doc))
        {
            row = SubTitle(ws, row + 1, $"{no++}. {table.Title}");
            row = ResultSheet(ws, row + 1, table) + 1;
        }
        foreach (ReviewReportTable t in doc.ResultTables)
        {
            row = SubTitle(ws, row + 1, $"{no++}. {FirstNonBlank(t.Caption, "Result")}");
            row = Table(ws, row, null, t.Columns, t.Rows) + 1;
        }
        row = Pictures(ws, row, picsBySection.GetValueOrDefault("results"), imageByNo);
        foreach (string f in doc.Findings.Where(f => !string.IsNullOrWhiteSpace(f)))
            row = Paragraph(ws, row, Dash(f));
        row++;

        row = Section(ws, row, "IV. Conclusion");
        row = Paragraph(ws, row, Dash(doc.Conclusion));
        foreach (string l in doc.Limitations.Where(l => !string.IsNullOrWhiteSpace(l)))
            row = Paragraph(ws, row, Dash(l));
        foreach (string n in doc.NextActions.Where(n => !string.IsNullOrWhiteSpace(n)))
            row = Paragraph(ws, row, "→ " + n.Trim());
        row = Pictures(ws, row + 1, picsBySection.GetValueOrDefault("conclusion"), imageByNo);

        if (picsBySection.TryGetValue("appendix", out var appendix) && appendix.Count > 0)
        {
            row = Section(ws, row + 1, "Appendix");
            row = Pictures(ws, row + 1, appendix, imageByNo);
        }
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.FitToPages(1, 0);

        // Second sheet: every result value as one row, so each number stays next to its metric and condition.
        var data = wb.Worksheets.Add("Data");
        data.Style.Font.FontName = FormFont;
        if (doc.Metrics.Count > 0)
        {
            DataSheet(data, doc);
            using var stream = new MemoryStream();
            wb.SaveAs(stream);
            return stream.ToArray();
        }
        string[] head = ["Table", "Metric", "Unit", "Condition", "Value"];
        for (int c = 0; c < head.Length; c++)
        {
            data.Cell(1, c + 1).Value = head[c];
            data.Cell(1, c + 1).Style.Font.Bold = true;
            data.Cell(1, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E2E8F0");
        }
        int dr = 2;
        foreach (ReviewReportTable t in doc.ResultTables)
        {
            // Tables follow ["Metric", "Unit", <conditions>...]; older ones have no Unit column.
            bool hasUnit = t.Columns.Count > 1 && t.Columns[1].Trim().Equals("Unit", StringComparison.OrdinalIgnoreCase);
            int firstValue = hasUnit ? 2 : 1;
            foreach (List<string> cells in t.Rows)
            {
                string metric = cells.FirstOrDefault() ?? "";
                string unit = hasUnit && cells.Count > 1 ? cells[1] : "";
                for (int c = firstValue; c < cells.Count; c++)
                {
                    if (string.IsNullOrWhiteSpace(cells[c])) continue;
                    data.Cell(dr, 1).Value = t.Caption;
                    data.Cell(dr, 2).Value = metric;
                    data.Cell(dr, 3).Value = unit;
                    data.Cell(dr, 4).Value = c < t.Columns.Count ? t.Columns[c] : $"Col {c + 1}";
                    SetCell(data.Cell(dr, 5), cells[c]);
                    dr++;
                }
            }
        }
        data.Columns(1, 5).AdjustToContents(5, 60);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static int Section(IXLWorksheet ws, int row, string title)
    {
        var range = ws.Range(row, 1, row, LastCol).Merge();
        range.Value = title;
        range.Style.Font.Bold = true;
        range.Style.Font.FontSize = 12;
        range.Style.Fill.BackgroundColor = SectionGrey;
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Row(row).Height = 28;
        return row + 1;
    }

    private static int SubTitle(IXLWorksheet ws, int row, string text)
    {
        ws.Cell(row, 1).Value = text;
        ws.Cell(row, 1).Style.Font.Bold = true;
        return row + 1;
    }

    private static string Label(string key, string? value) =>
        string.IsNullOrWhiteSpace(value) ? "" : $"{key}: {value.Trim()}";

    /// <summary>Prose lines in the form start with "- ", like the author's own report.</summary>
    private static string Dash(string? text)
    {
        string t = (text ?? "").Trim();
        return t.Length == 0 || t.StartsWith('-') || t.StartsWith('•') ? t : "- " + t;
    }

    private static int KeyValue(IXLWorksheet ws, int row, string key, string? value)
    {
        ws.Cell(row, 1).Value = key;
        ws.Cell(row, 1).Style.Font.Bold = true;
        if (!string.IsNullOrEmpty(value))
        {
            var range = ws.Range(row, 2, row, LastCol).Merge();
            range.Value = value;
            range.Style.Alignment.WrapText = true;
        }
        return row + 1;
    }

    private static int Paragraph(IXLWorksheet ws, int row, string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return row;
        var range = ws.Range(row, 1, row, LastCol).Merge();
        range.Value = text.Trim();
        range.Style.Alignment.WrapText = true;
        int lines = text.Trim().Split('\n').Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / 110.0)));
        ws.Row(row).Height = Math.Min(400, 15 * lines);
        return row + 1;
    }

    private static int Table(IXLWorksheet ws, int row, string? caption, List<string> columns, List<List<string>> rows)
    {
        if (columns.Count == 0 && rows.Count == 0) return row;
        if (!string.IsNullOrWhiteSpace(caption))
        {
            ws.Cell(row, 1).Value = caption;
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 1).Style.Font.Italic = true;
            row++;
        }
        int width = Math.Max(columns.Count, rows.Count == 0 ? 0 : rows.Max(r => r.Count));
        width = Math.Min(width, LastCol);
        int top = row;
        for (int c = 0; c < width; c++)
        {
            var cell = ws.Cell(row, c + 1);
            cell.Value = c < columns.Count ? columns[c] : "";
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#F1F5F9");
        }
        row++;
        foreach (List<string> cells in rows)
        {
            for (int c = 0; c < width; c++)
            {
                var cell = ws.Cell(row, c + 1);
                SetCell(cell, c < cells.Count ? cells[c] : "");
                cell.Style.Alignment.WrapText = true;
            }
            row++;
        }
        if (width > 0)
        {
            var box = ws.Range(top, 1, row - 1, width);
            box.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            box.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }
        return row;
    }

    /// <summary>
    /// Plain numbers go in as numbers (so the normalizer reads them from the cell), "5.7%" as a
    /// percent-formatted number that still shows 5.7%; everything else stays text.
    /// </summary>
    private static void SetCell(IXLCell cell, string value)
    {
        string v = (value ?? "").Trim();
        bool percent = v.EndsWith('%');
        string n = percent ? v[..^1].Trim() : v;
        bool codeLike = n.Length > 1 && n[0] == '0' && n[1] != '.';   // "007" is a code, not 7
        if (!codeLike && n.Length is > 0 and < 20
            && double.TryParse(n, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
        {
            if (percent)
            {
                int decimals = n.Contains('.') ? n.Length - n.IndexOf('.') - 1 : 0;
                cell.Value = d / 100.0;
                cell.Style.NumberFormat.Format = decimals == 0 ? "0%" : "0." + new string('0', decimals) + "%";
            }
            else
            {
                cell.Value = d;
            }
            return;
        }
        cell.Value = v;
    }

    private static void Box(IXLRange range, string? fill = null)
    {
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        if (fill is not null) range.Style.Fill.BackgroundColor = XLColor.FromHtml(fill);
    }

    /// <summary>
    /// One result table in the form: No | Date | (Model) | Type | metrics. An NG count has its NG rate
    /// in ppm (red) on the row under it, another count its paired rate; other cells span both rows.
    /// </summary>
    private static int ResultSheet(IXLWorksheet ws, int row, ReviewReportResultTable t)
    {
        var arms = t.Arms;
        var cols = t.Columns;
        int per = t.TwoRows ? 2 : 1;
        const int left = 2;

        var desc = new List<(string Head, Func<ReviewReportCondition, int, string> Get, bool Merge)>
        {
            ("No", (_, i) => (i + 1).ToString(CultureInfo.InvariantCulture), false),
        };
        if (arms.Any(a => !string.IsNullOrWhiteSpace(a.Date))) desc.Add(("Date", (a, _) => a.Date, true));
        if (arms.Any(a => !string.IsNullOrWhiteSpace(a.Model))) desc.Add(("Model", (a, _) => a.Model, true));
        desc.Add(("Type", (a, _) => a.Label, false));

        int width = desc.Count + cols.Count;
        int right = left + width - 1;
        int headRow = row;
        for (int c = 0; c < desc.Count; c++) ws.Cell(row, left + c).Value = desc[c].Head;
        for (int c = 0; c < cols.Count; c++)
            ws.Cell(row, left + desc.Count + c).Value = cols[c].HasBelow ? $"{cols[c].Header}\n/ {cols[c].BelowHeader}" : cols[c].Header;
        var head = ws.Range(row, left, row, right);
        head.Style.Font.Bold = true;
        head.Style.Alignment.WrapText = true;
        head.Style.Fill.BackgroundColor = HeaderBlue;
        // Tall enough for the longest wrapped header (about 10 characters per line at width 11).
        int lines = Enumerable.Range(left, width)
            .Select(c => ws.Cell(row, c).GetString().Split('\n').Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / 10.0))))
            .DefaultIfEmpty(1).Max();
        ws.Row(row).Height = Math.Max(30, lines * 15 + 6);
        row++;

        int top = row;
        int bottom = top + arms.Count * per - 1;
        for (int i = 0; i < arms.Count; i++)
        {
            int r = top + i * per;
            ReviewReportCondition arm = arms[i];
            string id = arm.Id.Trim();
            if (arm.Role == "test")
                ws.Range(r, left + desc.Count - 1, r + per - 1, right).Style.Fill.BackgroundColor = TestYellow;
            for (int c = 0; c < cols.Count; c++)
            {
                int col = left + desc.Count + c;
                ReviewReportColumn column = cols[c];
                SetCell(ws.Cell(r, col), column.Main.ValueOf(id));
                if (column.Main.Kind == "rate") ws.Cell(r, col).Style.Font.FontColor = RateRed;
                if (per == 2 && column.HasBelow)
                {
                    var below = ws.Cell(r + 1, col);
                    if (column.Ng && column.Ppm(id) is double ppm)
                    {
                        below.Value = ppm;
                        below.Style.NumberFormat.Format = "#,##0";
                    }
                    else
                    {
                        SetCell(below, column.Below(id));
                    }
                    below.Style.Font.FontColor = RateRed;
                    ws.Cell(r, col).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                }
                else if (per == 2)
                {
                    ws.Range(r, col, r + 1, col).Merge();
                }
            }
        }
        for (int d = 0; d < desc.Count; d++)
        {
            var values = arms.Select((a, i) => (desc[d].Get(a, i) ?? "").Trim()).ToList();
            int[] spans = desc[d].Merge ? Spans(values) : values.Select(_ => 1).ToArray();
            for (int i = 0; i < arms.Count; i++)
            {
                if (spans[i] == 0) continue;
                var range = ws.Range(top + i * per, left + d, top + (i + spans[i]) * per - 1, left + d);
                if (range.RowCount() > 1) range.Merge();
                range.FirstCell().Value = values[i];
                range.Style.Alignment.WrapText = true;
            }
        }
        var table = ws.Range(headRow, left, bottom, right);
        table.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        table.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        for (int c = left; c <= right; c++)
            ws.Range(headRow, c, bottom, c).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        ws.Range(headRow, left, headRow, right).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        for (int i = 1; i <= arms.Count; i++)
            ws.Range(top + i * per - 1, left, top + i * per - 1, right).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        for (int r = top; r <= bottom; r++) ws.Row(r).Height = 26;
        return bottom + 1;
    }

    /// <summary>The normalized results, one value per row, with the condition and metric fields beside it.</summary>
    private static void DataSheet(IXLWorksheet data, ReviewReportDoc doc)
    {
        string[] head = ["Arm ID", "Condition", "Role", "Date", "Model", "Metric ID", "Metric", "Unit", "Kind", "Rate of", "Value"];
        for (int c = 0; c < head.Length; c++) data.Cell(1, c + 1).Value = head[c];
        var h = data.Range(1, 1, 1, head.Length);
        h.Style.Font.Bold = true;
        h.Style.Fill.BackgroundColor = XLColor.FromHtml("#E2E8F0");
        var arms = ResultArms(doc).ToDictionary(a => a.Id.Trim(), StringComparer.OrdinalIgnoreCase);
        int r = 2;
        foreach (ReviewReportMetric m in doc.Metrics)
        {
            foreach (ReviewReportValue v in m.Values.Where(v => !string.IsNullOrWhiteSpace(v.Value)))
            {
                ReviewReportCondition? a = arms.GetValueOrDefault(v.ArmId.Trim());
                data.Cell(r, 1).Value = v.ArmId;
                data.Cell(r, 2).Value = a?.Label ?? "";
                data.Cell(r, 3).Value = a?.Role ?? "";
                data.Cell(r, 4).Value = a?.Date ?? "";
                data.Cell(r, 5).Value = a?.Model ?? "";
                data.Cell(r, 6).Value = m.Id;
                data.Cell(r, 7).Value = m.Name;
                data.Cell(r, 8).Value = m.Unit;
                data.Cell(r, 9).Value = m.Kind;
                data.Cell(r, 10).Value = m.RateOf;
                SetCell(data.Cell(r, 11), v.Value);
                r++;
            }
        }
        data.Columns(1, head.Length).AdjustToContents(5, 60);
    }

    /// <summary>
    /// Places the pictures of one section left to right from column B, each snapped to whole columns
    /// with a yellow label bar under it; a group of three or more gets one bar with the group name.
    /// Wraps to a new line when the row is full.
    /// </summary>
    private static int Pictures(IXLWorksheet ws, int row, List<ReviewReportImagePlacement>? placements,
        Dictionary<int, ReviewReportImage> images)
    {
        if (placements is null) return row;
        var groups = GroupPictures(placements).Select(g => g.Where(p => images.ContainsKey(p.No)).ToList())
            .Where(g => g.Count > 0).ToList();
        if (groups.Count == 0) return row;

        int height = groups.Sum(g => g.Count) == 1 ? 300 : 220;
        int picRows = (int)Math.Ceiling(height / PxPerRow);
        int lineTop = row, col = 2;
        bool lineUsed = false;
        foreach (List<ReviewReportImagePlacement> group in groups)
        {
            var spans = new List<(ReviewReportImagePlacement P, int Span, ClosedXML.Excel.Drawings.IXLPicture? Pic)>();
            foreach (ReviewReportImagePlacement p in group)
            {
                try
                {
                    var pic = ws.AddPicture(new MemoryStream(images[p.No].Data));
                    double w = pic.OriginalWidth * ((double)height / Math.Max(1, pic.OriginalHeight));
                    spans.Add((p, Math.Clamp((int)Math.Round(w / ColPx), 2, LastCol - 1), pic));
                }
                catch
                {
                    spans.Add((p, 3, null));
                }
            }
            int groupSpan = spans.Sum(x => x.Span);
            if (lineUsed && col + groupSpan - 1 > LastCol)
            {
                lineTop += picRows + 2;
                col = 2;
            }
            int groupStart = col;
            foreach (var (p, span, pic) in spans)
            {
                if (col + span - 1 > LastCol && col > 2)
                {
                    lineTop += picRows + 2;
                    col = groupStart = 2;
                }
                double w = span * ColPx - 6;
                if (pic is not null)
                {
                    double h = Math.Min(height, pic.OriginalHeight * (w / Math.Max(1, pic.OriginalWidth)));
                    pic.MoveTo(ws.Cell(lineTop, col), 3, (int)((height - h) / 2)).WithSize((int)w, (int)h);
                }
                else
                {
                    ws.Cell(lineTop, col).Value = $"(IMG {p.No} could not be embedded)";
                }
                if (group.Count < 3)
                    LabelBar(ws.Range(lineTop + picRows, col, lineTop + picRows, col + span - 1), FirstNonBlank(p.Caption, p.Group, $"IMG {p.No}"));
                col += span;
            }
            if (group.Count >= 3)
                LabelBar(ws.Range(lineTop + picRows, groupStart, lineTop + picRows, col - 1), FirstNonBlank(group[0].Group, group[0].Caption));
            lineUsed = true;
        }
        return lineTop + picRows + 2;
    }

    private static void LabelBar(IXLRange range, string text)
    {
        range.Merge();
        range.Value = text;
        range.Style.Font.Bold = true;
        range.Style.Font.FontName = FormFont;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        range.Style.Fill.BackgroundColor = TestYellow;
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.OutsideBorderColor = XLColor.FromHtml("#FF0000");
    }

    /// <summary>Groups placements by their group name in first-seen order; a picture without a group stands alone.</summary>
    public static List<List<ReviewReportImagePlacement>> GroupPictures(IEnumerable<ReviewReportImagePlacement> placements)
    {
        var groups = new List<List<ReviewReportImagePlacement>>();
        var byName = new Dictionary<string, List<ReviewReportImagePlacement>>(StringComparer.OrdinalIgnoreCase);
        foreach (ReviewReportImagePlacement p in placements)
        {
            string key = (p.Group ?? "").Trim();
            if (key.Length > 0 && byName.TryGetValue(key, out var existing))
            {
                existing.Add(p);
                continue;
            }
            var g = new List<ReviewReportImagePlacement> { p };
            groups.Add(g);
            if (key.Length > 0) byName[key] = g;
        }
        return groups;
    }

    // ── Readable views of the normalized results ────────────────────────────

    /// <summary>Conditions in the order listed, plus any condition id a metric uses that is not listed.</summary>
    public static List<ReviewReportCondition> ResultArms(ReviewReportDoc doc)
    {
        var arms = new List<ReviewReportCondition>(doc.Conditions);
        var known = arms.Select(a => a.Id.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (string id in doc.Metrics.SelectMany(m => m.Values).Select(v => v.ArmId.Trim()))
            if (id.Length > 0 && known.Add(id))
                arms.Add(new ReviewReportCondition { Id = id, Label = id });

        return arms;
    }

    /// <summary>
    /// Splits the results into the tables a person reads. Each review (condition "table") gets its own
    /// table; inside a review, conditions measured on the same items share one table (one row per
    /// condition), and conditions measured on different items get a table of their own. Every NG
    /// count carries its NG rate in ppm underneath; with two or more NG items a Total NG column is added.
    /// </summary>
    public static List<ReviewReportResultTable> ResultTables(ReviewReportDoc doc)
    {
        var tables = new List<ReviewReportResultTable>();
        var metrics = doc.Metrics.Where(m => !string.IsNullOrWhiteSpace(m.Id))
            .GroupBy(m => m.Id.Trim(), StringComparer.OrdinalIgnoreCase).Select(g => g.First()).ToList();
        if (metrics.Count == 0) return tables;

        // An NG rate is taken against the nearest input listed before it (a second test stage has its own input).
        var inputOf = new Dictionary<ReviewReportMetric, ReviewReportMetric?>();
        ReviewReportMetric? lastInput = metrics.FirstOrDefault(IsInput);
        foreach (ReviewReportMetric m in metrics)
        {
            if (IsInput(m)) lastInput = m;
            // A total NG is the whole lot's NG, so it is taken against the first input.
            inputOf[m] = IsTotal(m) ? metrics.FirstOrDefault(IsInput) : lastInput;
        }
        var rateOf = new Dictionary<ReviewReportMetric, ReviewReportMetric>();
        var attached = new HashSet<ReviewReportMetric>();
        void Attach(ReviewReportMetric main, ReviewReportMetric rate)
        {
            if (rateOf.ContainsKey(main) || attached.Contains(rate) || ReferenceEquals(main, rate)) return;
            rateOf[main] = rate;
            attached.Add(rate);
        }
        var rates = metrics.Where(IsRate).ToList();
        foreach (ReviewReportMetric r in rates.Where(r => !string.IsNullOrWhiteSpace(r.RateOf)))
            if (metrics.FirstOrDefault(m => m.Id.Trim().Equals(r.RateOf.Trim(), StringComparison.OrdinalIgnoreCase)) is { } main)
                Attach(main, r);
        foreach (ReviewReportMetric r in rates.Where(r => !attached.Contains(r)))
        {
            string bare = BareName(r.Name);
            if (metrics.FirstOrDefault(m => !IsRate(m) && BareName(m.Name).Equals(bare, StringComparison.OrdinalIgnoreCase)) is { } main)
                Attach(main, r);
            else if (IsTotalName(r.Name) && metrics.FirstOrDefault(IsTotal) is { } total)
                Attach(total, r);
        }
        var shown = metrics.Where(m => !attached.Contains(m)).ToList();

        var arms = ResultArms(doc);
        foreach (var review in arms.GroupBy(a => a.Table.Trim(), StringComparer.OrdinalIgnoreCase))
        {
            // Conditions measured on the same items read as one table; other item sets get their own.
            var parts = review
                .GroupBy(a => string.Join("|", shown.Where(m => m.ValueOf(a.Id.Trim()).Trim().Length > 0).Select(m => m.Id)))
                .ToList();
            foreach (var part in parts)
            {
                var partArms = part.ToList();
                bool Has(ReviewReportMetric m) => partArms.Any(a => m.ValueOf(a.Id.Trim()).Trim().Length > 0);
                ReviewReportMetric? InputFor(ReviewReportMetric m) => inputOf[m] is { } i && Has(i) ? i : null;
                var cols = shown.Where(Has)
                    .Select(m => new ReviewReportColumn(m, rateOf.GetValueOrDefault(m), InputFor(m), IsNg(m) || IsTotal(m)))
                    .ToList();

                var ngCols = cols.Where(c => IsNg(c.Main)).ToList();
                // Total NG only when every NG item counts against the same input.
                if (ngCols.Count >= 2 && !cols.Any(c => IsTotal(c.Main)) && ngCols.Select(c => c.Input).Distinct().Count() == 1)
                {
                    var total = new ReviewReportMetric { Id = "total_ng", Name = "Total NG", Unit = ngCols[0].Main.Unit, Kind = "total_ng" };
                    foreach (ReviewReportCondition a in partArms)
                    {
                        var counts = ngCols.Select(c => ParseNumber(c.Main.ValueOf(a.Id.Trim()))).ToList();
                        if (counts.All(n => n is not null))
                            total.Values.Add(new ReviewReportValue { ArmId = a.Id.Trim(), Value = counts.Sum(n => n!.Value).ToString("0.##", CultureInfo.InvariantCulture) });
                    }
                    // A source "Total NG rate" column now sits under the added Total NG.
                    ReviewReportColumn? totalRate = cols.FirstOrDefault(c => IsRate(c.Main) && IsTotalName(c.Main.Name));
                    if (totalRate is not null) cols.Remove(totalRate);
                    cols.Add(new ReviewReportColumn(total, totalRate?.Main, ngCols[0].Input, true));
                }
                string title = FirstNonBlank(review.Key, doc.ResultTitle, "Result");
                tables.Add(new ReviewReportResultTable(title, partArms, cols));
            }
        }
        return tables;
    }

    private static string Norm(string? s) => (s ?? "").Trim().ToLowerInvariant();
    private static bool IsInput(ReviewReportMetric m) =>
        Norm(m.Kind) == "input" || Norm(m.Id) is "input" or "input_qty" || Norm(m.Name) is "input" or "input qty" or "total" or "total qty" or "투입" or "투입수";
    private static bool IsOk(ReviewReportMetric m) =>
        Norm(m.Kind) == "ok" || Norm(m.Id) == "ok" || Norm(m.Name) is "ok" or "pass" or "양품";
    private static bool IsRate(ReviewReportMetric m) => Norm(m.Kind) == "rate";
    private static bool IsTotalName(string? name) =>
        (Norm(name).Contains("total") || Norm(name).Contains("합계") || Norm(name).Contains("총"))
        && (Norm(name).Contains("ng") || Norm(name).Contains("불량") || Norm(name).Contains("fail") || Norm(name).Contains("defect"));
    private static bool IsTotal(ReviewReportMetric m) =>
        Norm(m.Kind) == "total_ng" || (Norm(m.Kind) == "count" && IsTotalName(m.Name) && !IsInput(m) && !IsOk(m));
    /// <summary>An NG item: kind "ng", or (older reports) a count that is not input, OK or a total.</summary>
    private static bool IsNg(ReviewReportMetric m) =>
        Norm(m.Kind) == "ng" || (Norm(m.Kind) == "count" && !IsInput(m) && !IsOk(m) && !IsTotal(m));
    private static string BareName(string? name)
    {
        string n = Norm(name);
        foreach (string suffix in new[] { " rate", " ratio", " (%)", " %", "율", " 비율" })
            if (n.EndsWith(suffix)) n = n[..^suffix.Length].Trim();
        return n;
    }

    /// <summary>Rows that repeat the value above them are drawn as one merged cell; returns the span for each row (0 = covered).</summary>
    public static int[] Spans(IReadOnlyList<string> values)
    {
        var spans = new int[values.Count];
        for (int i = 0; i < values.Count;)
        {
            int j = i + 1;
            while (j < values.Count && values[j] == values[i] && values[i].Length > 0) j++;
            spans[i] = j - i;
            i = j;
        }
        return spans;
    }

    // ── Markdown (copy into a memo or mail) ─────────────────────────────────

    public static string ToMarkdown(ReviewReport report)
    {
        ReviewReportDoc d = report.Doc ?? new ReviewReportDoc();
        var sb = new StringBuilder();
        sb.AppendLine($"# {FirstNonBlank(d.Title, report.Title)}");
        sb.AppendLine();
        if (!string.IsNullOrWhiteSpace(d.ProductModel)) sb.AppendLine($"- Product model: {d.ProductModel}");
        if (!string.IsNullOrWhiteSpace(d.ReportDate)) sb.AppendLine($"- Report date: {d.ReportDate}");
        if (d.Tags.Count > 0) sb.AppendLine($"- Tags: {string.Join(", ", d.Tags)}");
        sb.AppendLine();
        sb.AppendLine("## I. Purpose").AppendLine(Dash(d.Purpose));
        if (!string.IsNullOrWhiteSpace(d.Background)) sb.AppendLine(Dash(d.Background));
        AppendPics(sb, d, "purpose");

        sb.AppendLine().AppendLine("## II. Content");
        foreach (string l in d.Content.Where(l => !string.IsNullOrWhiteSpace(l))) sb.AppendLine(Dash(l));
        foreach (var f in d.ChangedFactors.Where(f => !string.IsNullOrWhiteSpace(f.Name))) sb.AppendLine($"- {f.Name}: {f.Baseline} → {f.Changed}");
        foreach (var c in d.Conditions.Where(c => !string.IsNullOrWhiteSpace(c.Details))) sb.AppendLine($"- {c.Label}: {c.Details}");
        AppendPics(sb, d, "conditions");

        sb.AppendLine().AppendLine("## III. Result");
        int no = 1;
        foreach (ReviewReportResultTable table in ResultTables(d))
        {
            sb.AppendLine().AppendLine($"**{no++}. {table.Title}**").AppendLine();
            bool date = table.Arms.Any(a => !string.IsNullOrWhiteSpace(a.Date));
            var head = new List<string> { "No" };
            if (date) head.Add("Date");
            head.Add("Type");
            head.AddRange(table.Columns.Select(c => c.HasBelow ? $"{c.Header} / {c.BelowHeader}" : c.Header));
            AppendTable(sb, head, table.Arms.Select((a, i) =>
            {
                var cells = new List<string> { (i + 1).ToString(CultureInfo.InvariantCulture) };
                if (date) cells.Add(a.Date);
                cells.Add(a.Label);
                foreach (var c in table.Columns)
                {
                    string main = c.Main.ValueOf(a.Id.Trim());
                    string below = c.Below(a.Id.Trim());
                    cells.Add(below.Length > 0 ? $"{main} / {below}" : main);
                }
                return cells;
            }));
        }
        foreach (var t in d.ResultTables)
        {
            sb.AppendLine().AppendLine($"**{no++}. {FirstNonBlank(t.Caption, "Result")}**").AppendLine();
            AppendTable(sb, t.Columns, t.Rows);
        }
        AppendPics(sb, d, "results");
        foreach (string f in d.Findings.Where(f => !string.IsNullOrWhiteSpace(f))) sb.AppendLine(Dash(f));

        sb.AppendLine().AppendLine("## IV. Conclusion").AppendLine(Dash(d.Conclusion));
        foreach (string l in d.Limitations.Where(l => !string.IsNullOrWhiteSpace(l))) sb.AppendLine(Dash(l));
        foreach (string n in d.NextActions.Where(n => !string.IsNullOrWhiteSpace(n))) sb.AppendLine($"→ {n.Trim()}");
        AppendPics(sb, d, "conclusion");
        AppendPics(sb, d, "appendix");
        return sb.ToString();
    }

    private static void AppendPics(StringBuilder sb, ReviewReportDoc d, string section)
    {
        foreach (var g in GroupPictures(d.Images.Where(p => p.Section == section)))
        {
            string name = g[0].Group;
            sb.AppendLine().AppendLine((string.IsNullOrWhiteSpace(name) ? "" : $"*{name}*: ")
                + string.Join(" · ", g.Select(p => $"[IMG {p.No}] {p.Caption}")));
        }
    }

    private static void AppendTable(StringBuilder sb, IReadOnlyList<string> columns, IEnumerable<List<string>> rows)
    {
        static string Esc(string s) => (s ?? "").Replace("|", "\\|").Replace("\n", " ");
        sb.AppendLine("| " + string.Join(" | ", columns.Select(Esc)) + " |");
        sb.AppendLine("|" + string.Concat(columns.Select(_ => "---|")));
        foreach (var r in rows)
            sb.AppendLine("| " + string.Join(" | ", Enumerable.Range(0, columns.Count).Select(i => Esc(i < r.Count ? r[i] : ""))) + " |");
    }
}

// ── Models ──────────────────────────────────────────────────────────────────

public sealed record ReviewReportSummary(string Id, string Title, string CreatedBy, DateTime UpdatedAt);

public sealed record ReviewReportImage(int No, string FileName, string MediaType, byte[] Data)
{
    public string DataUrl => $"data:{MediaType};base64,{Convert.ToBase64String(Data)}";
}

public sealed class ReviewReport
{
    public string Id { get; set; } = "";
    public string Title { get; set; } = "";
    public string CreatedBy { get; set; } = "";
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }
    public string Language { get; set; } = ReviewReportService.Languages[0];
    public string InputText { get; set; } = "";
    public string ExtraInstruction { get; set; } = "";
    public ReviewReportDoc? Doc { get; set; }
    public string Model { get; set; } = "";
    public List<ReviewReportImage> Images { get; set; } = [];
    /// <summary>Change requests applied to the report after it was written, oldest first.</summary>
    public List<ReviewReportRevision> Revisions { get; set; } = [];
}

public sealed class ReviewReportRevision
{
    public DateTime At { get; set; }
    public string By { get; set; } = "";
    public string Request { get; set; } = "";
}

public sealed class ReviewReportDoc
{
    public string Title { get; set; } = "";
    public string ProductModel { get; set; } = "";
    public string ReportDate { get; set; } = "";
    public string Purpose { get; set; } = "";
    public string Background { get; set; } = "";
    /// <summary>"II. Content" lines: what was done.</summary>
    public List<string> Content { get; set; } = [];
    /// <summary>Heading over the result table ("Result check New Jig").</summary>
    public string ResultTitle { get; set; } = "";
    public List<ReviewReportFactor> ChangedFactors { get; set; } = [];
    public List<ReviewReportCondition> Conditions { get; set; } = [];
    /// <summary>Old reports only; new reports keep results in <see cref="Metrics"/>.</summary>
    public List<ReviewReportTable> ResultTables { get; set; } = [];
    public List<ReviewReportMetric> Metrics { get; set; } = [];
    public List<string> Findings { get; set; } = [];
    public string Conclusion { get; set; } = "";
    public List<string> Limitations { get; set; } = [];
    public List<string> NextActions { get; set; } = [];
    public List<string> Tags { get; set; } = [];
    public List<ReviewReportImagePlacement> Images { get; set; } = [];
}

public sealed class ReviewReportFactor
{
    public string Name { get; set; } = "";
    public string Baseline { get; set; } = "";
    public string Changed { get; set; } = "";
}

public sealed class ReviewReportCondition
{
    public string Id { get; set; } = "";
    public string Label { get; set; } = "";
    public string Role { get; set; } = "";
    public string Date { get; set; } = "";
    public string Model { get; set; } = "";
    public string Details { get; set; } = "";
    /// <summary>The review this condition belongs to; each review is its own result table.</summary>
    public string Table { get; set; } = "";
}

public sealed class ReviewReportTable
{
    public string Caption { get; set; } = "";
    public List<string> Columns { get; set; } = [];
    public List<List<string>> Rows { get; set; } = [];
    public string Source { get; set; } = "";
}

/// <summary>One measured item with one value per condition — the review-data metric shape.</summary>
public sealed class ReviewReportMetric
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Unit { get; set; } = "";
    public string Kind { get; set; } = "";
    /// <summary>For a rate: the id of the count metric it is the rate of.</summary>
    public string RateOf { get; set; } = "";
    public List<ReviewReportValue> Values { get; set; } = [];

    public string ValueOf(string armId) =>
        Values.FirstOrDefault(v => string.Equals(v.ArmId, armId, StringComparison.OrdinalIgnoreCase))?.Value ?? "";
}

public sealed class ReviewReportValue
{
    public string ArmId { get; set; } = "";
    public string Value { get; set; } = "";
}

/// <summary>
/// A result column for people. An NG column shows the count with its NG rate in ppm (no decimals)
/// underneath, from count / input; without an input it converts the source rate. Another count
/// with a rate shows that rate underneath as written.
/// </summary>
public sealed record ReviewReportColumn(ReviewReportMetric Main, ReviewReportMetric? Rate, ReviewReportMetric? Input, bool Ng)
{
    public string Header => string.IsNullOrWhiteSpace(Main.Unit) ? Main.Name : $"{Main.Name} ({Main.Unit})";
    public bool HasBelow => Ng || Rate is not null;
    public string BelowHeader => Ng ? "NG rate (ppm)" : Rate?.Name ?? "";

    public double? Ppm(string armId)
    {
        double? count = ReviewReportService.ParseNumber(Main.ValueOf(armId));
        double? input = Input is null ? null : ReviewReportService.ParseNumber(Input.ValueOf(armId));
        if (count is not null && input is > 0) return Math.Round(count.Value / input.Value * 1_000_000);
        double? rate = Rate is null ? null : ReviewReportService.ParseNumber(Rate.ValueOf(armId));
        if (rate is not null && Rate!.ValueOf(armId).Contains('%')) return Math.Round(rate.Value * 1_000_000);
        return null;
    }

    /// <summary>The text under the count: ppm for NG, else the paired rate as written.</summary>
    public string Below(string armId)
    {
        if (!Ng) return Rate?.ValueOf(armId) ?? "";
        return Ppm(armId) is double ppm ? ppm.ToString("#,##0", System.Globalization.CultureInfo.InvariantCulture) : "";
    }
}

/// <summary>One result table: a review (or one set of measured items within it), its conditions and columns.</summary>
public sealed record ReviewReportResultTable(string Title, List<ReviewReportCondition> Arms, List<ReviewReportColumn> Columns)
{
    public bool TwoRows => Columns.Any(c => c.HasBelow);
}

public sealed class ReviewReportImagePlacement
{
    public int No { get; set; }
    public string Caption { get; set; } = "";
    /// <summary>Pictures with the same group in a section are shown side by side.</summary>
    public string Group { get; set; } = "";
    public string Section { get; set; } = "appendix";
}
