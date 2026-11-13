using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace PortalGNLAB.Web.Services.GraphMaker;

public sealed record GraphMakerAiInput(string Text, string Delimiter, int HeaderRowNumber, bool HeaderNotUse);

/// <summary>The AI's pick: graph type index (client graphKinds order = <see cref="GraphMakerGraphType"/>), title, why, and the client GraphOptions keys it set.</summary>
public sealed record GraphMakerAiSuggestion(int Kind, string Title, string Reason, JsonObject Options);

public sealed record GraphMakerAiResult(int Kind, string Title, string Reason, JsonObject Options, string Model, double ElapsedSeconds);

/// <summary>
/// Working > Graphmaker: the user pastes a table and the Codex CLI reads a summary of it (columns, types, stats,
/// sample rows) and answers which Graph Maker chart fits and how to set it up. Nothing is saved; the client
/// renders the answer with the existing Graph Maker renderer and the user can still change every option.
/// </summary>
public sealed class GraphMakerAiService(AiProviderSettingsService settings, ILogger<GraphMakerAiService> logger)
{
    public const string Effort = "medium";
    public const string ModelLabel = "Codex CLI";
    private const int MaxSampleRows = 15;
    private const int MaxTailRows = 5;
    private const int MaxCellChars = 40;
    private static readonly TimeSpan Timeout = TimeSpan.FromMinutes(5);

    private static readonly string[] ColumnKeys =
    [
        "xColumn", "heatMapXColumn", "heatMapYColumn", "heatMapValueColumn", "heatMapNgColumn", "heatMapInputColumn",
        "graphSpecColumn", "graphUpperColumn", "graphLowerColumn",
    ];
    private static readonly string[] ColumnListKeys = ["xColumns", "yColumns"];
    private static readonly string[] BoolKeys =
    [
        "heatMapUseNgRate", "showProcessContour", "showProcessTrendLine", "showProcessMax", "showProcessMin", "showProcessAvg",
        "showProcessStd", "showProcessCpu", "showProcessCpl", "showProcessCpk", "showScatterMax", "showScatterMin",
        "showScatterAvg", "showScatterStd", "showScatterCpu", "showScatterCpl", "showScatterCpk",
    ];
    private static readonly string[] NumberKeys = ["upperLimit", "specLimit", "lowerLimit"];
    private static readonly string[] RowNumberKeys = ["processSpecRowNumber", "processUpperRowNumber", "processLowerRowNumber"];
    /// <summary>Two-value enums of the client options (0 or 1).</summary>
    private static readonly string[] SwitchKeys = ["heatMapRateUnit", "scatterPointLayout", "scatterLabelPosition", "lineMode"];

    public async Task<GraphMakerAiResult> AnalyzeAsync(GraphMakerAiInput input, string actor, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(input.Text)) throw new ArgumentException("Paste the data first.");
        if (!settings.CodexCliEnabled) throw new ArgumentException("Codex CLI is disabled in Setting.");
        GraphMakerTable table = GraphMakerParser.Parse(input.Text,
            new(string.IsNullOrWhiteSpace(input.Delimiter) ? "auto" : input.Delimiter, Math.Max(1, input.HeaderRowNumber), input.HeaderNotUse));
        if (!table.HasData) throw new ArgumentException("No rows found. Check the delimiter and the header row.");

        var watch = Stopwatch.StartNew();
        string answer = await RunCodexAsync(BuildPrompt(table), actor, ct);
        GraphMakerAiSuggestion s = ParseAnswer(answer, table);
        return new GraphMakerAiResult(s.Kind, s.Title, s.Reason, s.Options, ModelLabel, Math.Round(watch.Elapsed.TotalSeconds, 1));
    }

    /// <summary>Instructions plus a compact picture of the table: each column's type and stats, then the first and last rows.</summary>
    public static string BuildPrompt(GraphMakerTable table)
    {
        var text = new StringBuilder();
        text.AppendLine(Instructions);
        text.AppendLine();
        text.AppendLine($"=== TABLE: {table.Rows.Count} data rows, {table.Headers.Count} columns ===");
        text.AppendLine("Columns (exact names, use them verbatim):");
        for (int c = 0; c < table.Headers.Count; c++) text.AppendLine($"- {DescribeColumn(table, c)}");
        text.AppendLine();
        text.AppendLine($"First rows (data row numbers from 1, tab separated):");
        text.AppendLine(string.Join('\t', table.Headers.Select(Clip)));
        int head = Math.Min(MaxSampleRows, table.Rows.Count);
        for (int r = 0; r < head; r++) text.AppendLine($"{r + 1}: {Row(table.Rows[r])}");
        int tailStart = Math.Max(head, table.Rows.Count - MaxTailRows);
        if (tailStart > head) text.AppendLine("...");
        for (int r = tailStart; r < table.Rows.Count; r++) text.AppendLine($"{r + 1}: {Row(table.Rows[r])}");
        return text.ToString();

        static string Row(string[] cells) => string.Join('\t', cells.Select(Clip));
    }

    private static string DescribeColumn(GraphMakerTable table, int index)
    {
        var values = table.Rows.Select(r => table.GetCell(r, index)).Where(v => v.Length > 0).ToList();
        var numbers = values.Select(v => GraphMakerPayloadBuilder.ParseNullable(v.Replace(",", ""))).OfType<double>().Where(double.IsFinite).ToList();
        int distinct = values.Distinct(StringComparer.Ordinal).Count();
        string name = $"\"{table.Headers[index]}\": {values.Count}/{table.Rows.Count} filled, {distinct} distinct";
        if (values.Count > 0 && numbers.Count == values.Count)
            return $"{name}, numeric min {N(numbers.Min())} max {N(numbers.Max())} mean {N(numbers.Average())}";
        string samples = string.Join(", ", values.Distinct(StringComparer.Ordinal).Take(5).Select(v => "\"" + Clip(v) + "\""));
        string kind = numbers.Count > 0 ? $"mixed text/numeric ({numbers.Count} numeric)" : "text";
        return $"{name}, {kind}, e.g. {samples}";

        static string N(double v) => v.ToString("G6", CultureInfo.InvariantCulture);
    }

    private static string Clip(string value)
    {
        string v = (value ?? "").Replace('\t', ' ').Replace('\n', ' ');
        return v.Length <= MaxCellChars ? v : v[..MaxCellChars] + "…";
    }

    /// <summary>
    /// Reads the AI answer leniently (code fences, text around the JSON) and keeps only what the renderer can use:
    /// a known graph type, real column names (matched case-insensitively and returned with the table's spelling),
    /// and option values of the right type. Anything else is dropped.
    /// </summary>
    public static GraphMakerAiSuggestion ParseAnswer(string? text, GraphMakerTable table)
    {
        const string fail = "The AI answer could not be read. Try again.";
        string s = text ?? "";
        int a = s.IndexOf('{'), b = s.LastIndexOf('}');
        if (a < 0 || b <= a) throw new ArgumentException(fail);
        JsonObject root;
        try
        {
            root = JsonNode.Parse(s[a..(b + 1)]) as JsonObject ?? throw new ArgumentException(fail);
        }
        catch (JsonException) { throw new ArgumentException(fail); }

        int kind = ReadKind(root["kind"]) ?? ReadKind(root["graphType"]) ?? throw new ArgumentException(fail);
        string title = Str(root["title"]);
        string reason = Str(root["reason"]);
        JsonObject options = NormalizeOptions(root["options"] as JsonObject, table);
        return new GraphMakerAiSuggestion(kind, title, reason, options);
    }

    public static JsonObject NormalizeOptions(JsonObject? raw, GraphMakerTable table)
    {
        var result = new JsonObject();
        if (raw is null) return result;
        var lookup = new Dictionary<string, JsonNode?>(StringComparer.OrdinalIgnoreCase);
        foreach ((string key, JsonNode? value) in raw) lookup[key] = value;

        foreach (string key in ColumnKeys)
            if (lookup.TryGetValue(key, out JsonNode? v) && Column(table, v) is { } column) result[key] = column;

        foreach (string key in ColumnListKeys)
        {
            if (!lookup.TryGetValue(key, out JsonNode? v) || v is not JsonArray list) continue;
            var columns = list.Select(x => Column(table, x)).OfType<string>().Distinct(StringComparer.Ordinal).ToList();
            if (columns.Count > 0) result[key] = new JsonArray(columns.Select(c => (JsonNode?)JsonValue.Create(c)).ToArray());
        }

        foreach (string key in BoolKeys)
            if (lookup.TryGetValue(key, out JsonNode? v) && v is JsonValue jv && jv.TryGetValue(out bool flag)) result[key] = flag;

        foreach (string key in NumberKeys)
            if (lookup.TryGetValue(key, out JsonNode? v) && Number(v) is { } n) result[key] = n;

        foreach (string key in RowNumberKeys)
            if (lookup.TryGetValue(key, out JsonNode? v) && Number(v) is { } n && n >= 1 && n <= table.Rows.Count && n == Math.Floor(n))
                result[key] = (int)n;

        foreach (string key in SwitchKeys)
            if (lookup.TryGetValue(key, out JsonNode? v) && Number(v) is { } n && (n == 0 || n == 1)) result[key] = (int)n;

        if (lookup.TryGetValue("graphColumnGroups", out JsonNode? g) && g is JsonObject groups)
        {
            var clean = new JsonObject();
            foreach ((string column, JsonNode? group) in groups)
                if (Column(table, JsonValue.Create(column)) is { } name && Str(group) is { Length: > 0 } label) clean[name] = label;
            if (clean.Count > 0) result["graphColumnGroups"] = clean;
        }
        return result;
    }

    private static readonly string[][] KindNames =
    [
        ["line", "graph"], ["scatter"], ["average"], ["cpk"], ["normaldistribution", "normal", "distribution"],
        ["noxmultiy"], ["processtrend"], ["heatmap"],
    ];

    private static int? ReadKind(JsonNode? node)
    {
        if (node is not JsonValue value) return null;
        if (value.TryGetValue(out int i)) return i is >= 0 and <= 7 ? i : null;
        if (value.TryGetValue(out double d)) return d is >= 0 and <= 7 && d == Math.Floor(d) ? (int)d : null;
        if (!value.TryGetValue(out string? s) || s is null) return null;
        if (int.TryParse(s.Trim(), out int n)) return n is >= 0 and <= 7 ? n : null;
        string key = new(s.Where(char.IsLetter).Select(char.ToLowerInvariant).ToArray());
        for (int k = 0; k < KindNames.Length; k++)
            if (KindNames[k].Contains(key)) return k;
        return null;
    }

    private static string? Column(GraphMakerTable table, JsonNode? node)
    {
        string name = Str(node);
        if (name.Length == 0) return null;
        return table.Headers.FirstOrDefault(h => h == name)
            ?? table.Headers.FirstOrDefault(h => string.Equals(h.Trim(), name, StringComparison.OrdinalIgnoreCase));
    }

    private static double? Number(JsonNode? node)
    {
        if (node is not JsonValue value) return null;
        if (value.TryGetValue(out double d)) return double.IsFinite(d) ? d : null;
        if (value.TryGetValue(out string? s)
            && double.TryParse((s ?? "").Replace(",", "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double p))
            return double.IsFinite(p) ? p : null;
        return null;
    }

    private static string Str(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue(out string? s) ? (s ?? "").Trim() : "";

    /// <summary>Runs <c>codex exec</c> read-only in an empty temp folder and returns its last message.</summary>
    private async Task<string> RunCodexAsync(string prompt, string actor, CancellationToken ct)
    {
        string tmpDir = Path.Combine(Path.GetTempPath(), "portal-graph-ai", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmpDir);
        string promptPath = Path.Combine(tmpDir, "prompt.txt");
        string outputPath = Path.Combine(tmpDir, "answer.txt");
        string scriptPath = Path.Combine(tmpDir, "run.ps1");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(Timeout);
        try
        {
            var utf8 = new UTF8Encoding(false);
            await File.WriteAllTextAsync(promptPath, prompt, utf8, ct);
            await File.WriteAllTextAsync(scriptPath, BuildScript(tmpDir, promptPath, outputPath), utf8, ct);

            var psi = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                WorkingDirectory = tmpDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
            };
            foreach (string arg in new[] { "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath }) psi.ArgumentList.Add(arg);

            using var proc = new Process { StartInfo = psi };
            logger.LogInformation("Graph AI start [{Actor}] promptChars={Chars}", actor, prompt.Length);
            proc.Start();
            ChildProcessJob.Assign(proc);
            Task<string> stdout = proc.StandardOutput.ReadToEndAsync(timeout.Token);
            Task<string> stderr = proc.StandardError.ReadToEndAsync(timeout.Token);
            using (timeout.Token.Register(() => { try { if (!proc.HasExited) proc.Kill(entireProcessTree: true); } catch { } }))
            {
                try { await proc.WaitForExitAsync(timeout.Token); }
                catch (OperationCanceledException) when (!ct.IsCancellationRequested)
                {
                    throw new InvalidOperationException($"The AI did not answer within {Timeout.TotalMinutes:0} minutes.");
                }
            }
            string err = await stderr;
            await stdout;
            logger.LogInformation("Graph AI exit [{Actor}] code={Code}", actor, proc.ExitCode);
            string answer = File.Exists(outputPath) ? await File.ReadAllTextAsync(outputPath, ct) : "";
            if (proc.ExitCode != 0 || answer.Trim().Length == 0)
            {
                logger.LogWarning("Graph AI failed [{Actor}] code={Code}: {Err}", actor, proc.ExitCode, err.Length > 2000 ? err[^2000..] : err);
                throw new InvalidOperationException($"Codex CLI failed (exit {proc.ExitCode}). Check that the codex CLI is installed and logged in on the server.");
            }
            return answer;
        }
        finally
        {
            try { Directory.Delete(tmpDir, recursive: true); } catch { }
        }
    }

    public static string BuildScript(string workDir, string promptPath, string outputPath)
    {
        static string Q(string value) => "'" + value.Replace("'", "''") + "'";
        return string.Join(Environment.NewLine,
            "$ErrorActionPreference = 'Stop'",
            "$utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false",
            "[Console]::InputEncoding = $utf8NoBom",
            "[Console]::OutputEncoding = $utf8NoBom",
            "$OutputEncoding = $utf8NoBom",
            $"$prompt = Get-Content -LiteralPath {Q(promptPath)} -Raw -Encoding UTF8",
            $"$prompt | & codex exec -s read-only --ephemeral --skip-git-repo-check --color never -c 'model_reasoning_effort=\"{Effort}\"' --cd {Q(workDir)} -o {Q(outputPath)} -",
            "exit $LASTEXITCODE");
    }

    private const string Instructions = """
        You pick the chart for a table that a factory engineer pasted into the portal's Graph Maker. Do not run any
        command or read any file; everything you need is below. Look at what the columns mean (names, units, value
        ranges, text vs numbers, row order) and choose the ONE graph type that shows the data best, then set it up.

        Graph types (kind = number):
        0 Line ("Graph"): each Y column as a line over the rows; xColumn gives the x labels (dates, sample numbers).
          Columns named REF/UP/LOW can be set as graphSpecColumn/graphUpperColumn/graphLowerColumn (drawn as limits).
          graphColumnGroups maps a Y column to a group name; columns of one group share a colour (e.g. "Gap#1","Gap#2" -> "Gap").
          lineMode 0 = straight, 1 = smoothed.
        1 Scatter: points of Y columns. xColumns = category/X columns (one numeric X = true x/y scatter; text X or several
          X columns = points grouped by category). No xColumns = one category per Y column (compare distributions of several
          measurements). upperLimit/specLimit/lowerLimit draw spec lines. showScatterMax/Min/Avg/Std/Cpu/Cpl/Cpk add stats.
          scatterPointLayout 0 = aligned, 1 = spread (jitter, good for many repeated values).
        2 Average: average of each Y column per distinct xColumn value (e.g. per day, per line, per model).
        3 CPK: rolling Cpk of each Y column; needs upperLimit and/or lowerLimit.
        4 Normal distribution: histogram-like normal curve of each Y column (single measurement spread); limits optional.
        5 No X multi Y: each row is one line across the Y columns (profiles, e.g. a frequency response per sample where the
          columns are frequencies or positions).
        6 Process trend: xColumn is the input/process parameter, each Y column one chart of its effect (x/y scatter with
          optional trend line and stats). Rows holding the limits can be given as processUpperRowNumber /
          processLowerRowNumber / processSpecRowNumber (data row numbers as shown below). showProcessTrendLine,
          showProcessContour, showProcessMax/Min/Avg/Std/Cpu/Cpl/Cpk.
        7 Heatmap: heatMapXColumn x heatMapYColumn grid coloured by heatMapValueColumn, or by NG rate when
          heatMapUseNgRate is true (heatMapNgColumn / heatMapInputColumn, heatMapRateUnit 0 = %, 1 = ppm).

        Answer ONLY one JSON object, no other text:
        {"kind":<0-7>,"title":"<short chart title>","reason":"<one or two English sentences: why this chart fits this data>",
         "options":{ ...only the keys you want to set... }}
        Option keys: xColumn, xColumns, yColumns, heatMapXColumn, heatMapYColumn, heatMapValueColumn, heatMapUseNgRate,
        heatMapNgColumn, heatMapInputColumn, heatMapRateUnit, graphSpecColumn, graphUpperColumn, graphLowerColumn,
        graphColumnGroups, upperLimit, specLimit, lowerLimit, processSpecRowNumber, processUpperRowNumber,
        processLowerRowNumber, showProcessContour, showProcessTrendLine, showProcessMax, showProcessMin, showProcessAvg,
        showProcessStd, showProcessCpu, showProcessCpl, showProcessCpk, scatterPointLayout, scatterLabelPosition,
        lineMode, showScatterMax, showScatterMin, showScatterAvg, showScatterStd, showScatterCpu, showScatterCpl, showScatterCpk.
        Rules:
        - Column values must be exact column names from the list. yColumns must be numeric columns.
        - Do not put an ID/serial/date/limit column in yColumns. Leave out options you do not need.
        - Only set limits you can read from the data (limit columns or rows, or names like USL/LSL); never invent them.
        """;
}
