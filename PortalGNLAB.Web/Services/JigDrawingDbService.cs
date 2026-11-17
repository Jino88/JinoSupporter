using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace PortalGNLAB.Web.Services;

public sealed record JigDrawingMatch(string DrawingNo, string Name, string Model, string Category, int Latest);

/// <summary>One drawing number checked; <paramref name="Versions"/> = every revision of its base in the DB (-00, -01 …), oldest first.</summary>
public sealed record JigDrawingCheckItem(string DrawingNo, string Status, JigDrawingMatch? Match, List<JigDrawingMatch> Versions);

public sealed record JigDrawingCheckResult(bool Ok, string Error, string Path, List<JigDrawingCheckItem> Results);

public sealed record JigDrawingSuggestion(string DrawingNo, string Name, string Model, string Category, int Assy, int Latest, double Score);

public sealed record JigDrawingSuggestResult(bool Ok, string Error, string Path, List<JigDrawingSuggestion> Suggestions);

public sealed record JigDrawingDbView(string Path, string DefaultPath, string Resolved);

/// <summary>
/// Read-only lookup in the JIG/MC drawing DB (ORIGINAL_DRAWING.db, made by another tool): confirms the Drawing No
/// of a Working Request and suggests drawing numbers from a JIG name. The DB is only ever opened Mode=ReadOnly
/// and no index is created; drawing_no has none, so each query is one full scan.
/// </summary>
public sealed class JigDrawingDbService
{
    public const string SettingKey = "Working:JigMcDrawingDb";
    public const string DefaultPath = @"E:\db\JigMcDrawingDB";
    public const string DbFileName = "ORIGINAL_DRAWING.db";
    public const int MaxCheck = 1000;

    private readonly Func<string?> _getSetting;
    private readonly Action<string> _setSetting;

    public JigDrawingDbService(NgRateSettingsService settings)
        : this(() => settings.GetSetting(SettingKey), v => settings.SetSetting(SettingKey, v)) { }

    public JigDrawingDbService(Func<string?> getSetting, Action<string> setSetting)
    {
        _getSetting = getSetting;
        _setSetting = setSetting;
    }

    public string ConfiguredPath => _getSetting() is { Length: > 0 } s ? s : DefaultPath;

    public void SavePath(string? path) => _setSetting((path ?? string.Empty).Trim());

    /// <summary>Folder → ORIGINAL_DRAWING.db in it (else the first *.db); file → that file. "" when nothing is there.</summary>
    public static string Resolve(string? path)
    {
        string p = (path ?? string.Empty).Trim().Trim('"');
        if (p.Length == 0) return "";
        try
        {
            if (File.Exists(p)) return p;
            if (!Directory.Exists(p)) return "";
            string named = System.IO.Path.Combine(p, DbFileName);
            if (File.Exists(named)) return named;
            return Directory.EnumerateFiles(p, "*.db").OrderBy(f => f, StringComparer.OrdinalIgnoreCase).FirstOrDefault() ?? "";
        }
        catch { return ""; }
    }

    public JigDrawingDbView View() => new(ConfiguredPath, DefaultPath, Resolve(ConfiguredPath));

    private static SqliteConnection OpenReadOnly(string file)
    {
        var conn = new SqliteConnection(new SqliteConnectionStringBuilder
        {
            DataSource = file, Mode = SqliteOpenMode.ReadOnly, Pooling = false,
        }.ToString());
        conn.Open();
        return conn;
    }

    /// <summary>Row count and model count, for the Setting page's test button.</summary>
    public (bool Ok, string Message) Test()
    {
        string file = Resolve(ConfiguredPath);
        if (file.Length == 0) return (false, $"Drawing DB not found: {ConfiguredPath}");
        try
        {
            using SqliteConnection conn = OpenReadOnly(file);
            using SqliteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT COUNT(*), COUNT(DISTINCT model) FROM drawings;";
            using SqliteDataReader r = cmd.ExecuteReader();
            r.Read();
            return (true, $"{r.GetInt64(0):N0} drawings · {r.GetInt64(1):N0} models");
        }
        catch (Exception ex) { return (false, ex.Message); }
    }

    // ── check ───────────────────────────────────────────────────────────────────

    public JigDrawingCheckResult Check(IEnumerable<string>? drawingNos)
    {
        string file = Resolve(ConfiguredPath);
        if (file.Length == 0)
            return new(false, $"Drawing DB not found: {ConfiguredPath}", "", []);
        var requested = (drawingNos ?? []).Select(d => (d ?? string.Empty).Trim()).Take(MaxCheck).ToList();
        try
        {
            var keys = requested.Select(d => d.ToUpperInvariant()).Where(k => k.Length > 0).Distinct().ToList();
            var exact = new Dictionary<string, JigDrawingMatch>();
            var byBase = new Dictionary<string, List<(string Base, JigDrawingMatch M)>>();
            if (keys.Count > 0)
            {
                using SqliteConnection conn = OpenReadOnly(file);
                using SqliteCommand cmd = conn.CreateCommand();
                // one scan: exact numbers, plus anything sharing the base of a requested number
                var ps = new List<string>();
                for (int i = 0; i < keys.Count; i++)
                {
                    cmd.Parameters.AddWithValue("@d" + i, keys[i]);
                    cmd.Parameters.AddWithValue("@b" + i, BaseOf(keys[i]));
                    ps.Add("@d" + i);
                }
                string dIn = string.Join(",", ps);
                string bIn = string.Join(",", ps.Select(p => "@b" + p[2..]));
                cmd.CommandText = $"""
                    SELECT drawing_no, name, model, category, latest, base FROM drawings
                    WHERE UPPER(drawing_no) IN ({dIn}) OR UPPER(base) IN ({bIn});
                    """;
                using SqliteDataReader r = cmd.ExecuteReader();
                while (r.Read())
                {
                    var m = new JigDrawingMatch(Str(r, 0), Str(r, 1), Str(r, 2), Str(r, 3), r.IsDBNull(4) ? 0 : r.GetInt32(4));
                    exact[m.DrawingNo.ToUpperInvariant()] = m;
                    string b = Str(r, 5).ToUpperInvariant();
                    if (!byBase.TryGetValue(b, out var list)) byBase[b] = list = [];
                    list.Add((b, m));
                }
            }

            var results = new List<JigDrawingCheckItem>();
            foreach (string req in requested)
            {
                string key = req.ToUpperInvariant();
                if (key.Length == 0) { results.Add(new(req, "missing", null, [])); continue; }
                List<JigDrawingMatch> versions = byBase.TryGetValue(BaseOf(key), out var sibs)
                    ? sibs.Select(s => s.M).DistinctBy(m => m.DrawingNo.ToUpperInvariant())
                        .OrderBy(m => m.DrawingNo, StringComparer.OrdinalIgnoreCase).ToList()
                    : [];
                if (exact.TryGetValue(key, out var hit)) { results.Add(new(req, "found", hit, versions)); continue; }
                if (sibs is { Count: > 0 })
                {
                    JigDrawingMatch best = sibs.Select(s => s.M)
                        .OrderByDescending(m => m.Latest == 1)
                        .ThenByDescending(m => m.DrawingNo, StringComparer.OrdinalIgnoreCase)
                        .First();
                    results.Add(new(req, "revision", best, versions));
                    continue;
                }
                results.Add(new(req, "missing", null, []));
            }
            return new(true, "", file, results);
        }
        catch (Exception ex) { return new(false, ex.Message, file, []); }
    }

    private static readonly Regex RevisionTail = new(@"-\d{2}$", RegexOptions.Compiled);

    /// <summary>Drawing number without the trailing revision (-00).</summary>
    public static string BaseOf(string drawingNo) => RevisionTail.Replace(drawingNo.Trim(), "");

    private static string Str(SqliteDataReader r, int i) => r.IsDBNull(i) ? "" : r.GetString(i);

    // ── suggest ─────────────────────────────────────────────────────────────────

    private static readonly Regex LeadingParen = new(@"^\s*(\([^)]*\)\s*)+", RegexOptions.Compiled);
    private static readonly Regex TokenSplit = new(@"[^A-Z0-9]+", RegexOptions.Compiled);

    private static List<string> Tokens(string name) =>
        TokenSplit.Split(LeadingParen.Replace(name ?? string.Empty, "").ToUpperInvariant())
            .Where(t => t.Length > 0).Distinct().ToList();

    public JigDrawingSuggestResult Suggest(string? name, string? model, int limit = 10)
    {
        string file = Resolve(ConfiguredPath);
        if (file.Length == 0)
            return new(false, $"Drawing DB not found: {ConfiguredPath}", "", []);
        limit = limit <= 0 ? 10 : Math.Min(limit, 30);
        var tokens = Tokens(name ?? "");
        if (tokens.Count == 0) return new(true, "", file, []);
        string wantModel = (model ?? "").Trim().ToUpperInvariant();
        bool wantsJig = (name ?? "").Contains("JIG", StringComparison.OrdinalIgnoreCase);
        try
        {
            // a drawing number (or part of one) typed: the drawings whose number has it, of any model
            var byNo = SuggestByNumber(file, (name ?? "").Trim(), limit);
            if (byNo.Count > 0) return new(true, "", file, byNo);

            var rows = new List<(string No, string Base, string Name, string Model, string Cat, int Assy, int Latest)>();
            using (SqliteConnection conn = OpenReadOnly(file))
            using (SqliteCommand cmd = conn.CreateCommand())
            {
                var likes = new List<string>();
                for (int i = 0; i < tokens.Count; i++)
                {
                    cmd.Parameters.AddWithValue("@t" + i, "%" + tokens[i].Replace("\\", "\\\\").Replace("%", "\\%").Replace("_", "\\_") + "%");
                    likes.Add($"UPPER(name) LIKE @t{i} ESCAPE '\\'");
                }
                // a model given: only the drawings of that model
                string ofModel = "";
                if (wantModel.Length > 0)
                {
                    cmd.Parameters.AddWithValue("@m", wantModel);
                    ofModel = " AND UPPER(TRIM(model)) = @m";
                }
                cmd.CommandText = $"""
                    SELECT drawing_no, base, name, model, category, assy, latest FROM drawings
                    WHERE ({string.Join(" OR ", likes)}){ofModel} LIMIT 20000;
                    """;
                using SqliteDataReader r = cmd.ExecuteReader();
                while (r.Read())
                    rows.Add((Str(r, 0), Str(r, 1), Str(r, 2), Str(r, 3), Str(r, 4),
                        r.IsDBNull(5) ? 0 : r.GetInt32(5), r.IsDBNull(6) ? 0 : r.GetInt32(6)));
            }

            var scored = rows.Select(x =>
            {
                var rowTokens = Tokens(x.Name).ToHashSet();
                double score = (double)tokens.Count(rowTokens.Contains) / tokens.Count * 100;
                if (rowTokens.Count > 0) score -= Math.Min(10, (rowTokens.Count - tokens.Count(rowTokens.Contains)) * 1.0); // extra words
                if (wantsJig && x.Cat == "JIG") score += 10;
                if (x.Latest == 1) score += 5;
                if (x.Assy == 1) score += 3;
                return (Row: x, Score: score);
            })
            .Where(s => s.Score > 0)
            .GroupBy(s => string.IsNullOrEmpty(s.Row.Base) ? BaseOf(s.Row.No).ToUpperInvariant() : s.Row.Base.ToUpperInvariant())
            .Select(g => g.OrderByDescending(s => s.Row.Latest).ThenByDescending(s => s.Score)
                .ThenByDescending(s => s.Row.No, StringComparer.OrdinalIgnoreCase).First())
            .OrderByDescending(s => s.Score).ThenBy(s => s.Row.No, StringComparer.OrdinalIgnoreCase)
            .Take(limit)
            .Select(s => new JigDrawingSuggestion(s.Row.No, s.Row.Name, s.Row.Model, s.Row.Cat, s.Row.Assy, s.Row.Latest,
                Math.Round(s.Score, 1)))
            .ToList();
            return new(true, "", file, scored);
        }
        catch (Exception ex) { return new(false, ex.Message, file, []); }
    }

    /// <summary>
    /// The drawings whose number has the text typed (one word with a digit, e.g. BJJ2610-SP-58), the exact one first, then
    /// those starting with it, latest first. A drawing number is the same in every model, so the model is not asked.
    /// </summary>
    private static List<JigDrawingSuggestion> SuggestByNumber(string file, string text, int limit)
    {
        string q = text.ToUpperInvariant();
        if (q.Length < 3 || q.Any(char.IsWhiteSpace) || !q.Any(char.IsDigit)) return [];
        var rows = new List<JigDrawingSuggestion>();
        using SqliteConnection conn = OpenReadOnly(file);
        using SqliteCommand cmd = conn.CreateCommand();
        cmd.Parameters.AddWithValue("@q", "%" + q.Replace("\\", "\\\\").Replace("%", "\\%").Replace("_", "\\_") + "%");
        cmd.CommandText = """
            SELECT drawing_no, name, model, category, assy, latest FROM drawings
            WHERE UPPER(drawing_no) LIKE @q ESCAPE '\' LIMIT 2000;
            """;
        using SqliteDataReader r = cmd.ExecuteReader();
        while (r.Read())
        {
            string no = Str(r, 0);
            string up = no.ToUpperInvariant();
            double score = up == q ? 100 : up.StartsWith(q, StringComparison.Ordinal) ? 90 : 80;
            rows.Add(new(no, Str(r, 1), Str(r, 2), Str(r, 3), r.IsDBNull(4) ? 0 : r.GetInt32(4), r.IsDBNull(5) ? 0 : r.GetInt32(5), score));
        }
        return rows.DistinctBy(s => s.DrawingNo.ToUpperInvariant())
            .OrderByDescending(s => s.Score).ThenByDescending(s => s.Latest)
            .ThenBy(s => s.DrawingNo, StringComparer.OrdinalIgnoreCase)
            .Take(limit).ToList();
    }
}
