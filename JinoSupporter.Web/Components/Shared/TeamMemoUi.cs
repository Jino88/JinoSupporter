using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using JinoSupporter.Web.Services;

namespace JinoSupporter.Web.Components.Shared;

/// <summary>Small display helpers shared by the team memo pages.</summary>
public static class TeamMemoUi
{
    private static readonly string[] AvatarColors =
        ["#2563eb", "#db2777", "#059669", "#d97706", "#7c3aed", "#0891b2", "#dc2626", "#4f46e5"];

    public static string Initials(string name)
    {
        name = (name ?? "").Trim();
        if (name.Length == 0) return "?";
        string[] parts = name.Split([' ', '.', '_', '-'], StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 2) return (char.ToUpperInvariant(parts[0][0]).ToString() + char.ToUpperInvariant(parts[1][0])).ToString();
        return name.Length >= 2 ? name[..2].ToUpperInvariant() : name.ToUpperInvariant();
    }

    /// <summary>Stable colour per user name so the same person keeps the same avatar everywhere.</summary>
    public static string AvatarColor(string name)
    {
        int h = 17;
        foreach (char c in (name ?? "").ToLowerInvariant()) h = unchecked(h * 31 + c);
        return AvatarColors[(h & 0x7fffffff) % AvatarColors.Length];
    }

    public static string StatusClass(string status) => status switch
    {
        TeamMemoStatus.Done       => "tm-pill-done",
        TeamMemoStatus.InProgress => "tm-pill-prog",
        TeamMemoStatus.Hold       => "tm-pill-hold",
        _                         => "tm-pill-todo",
    };

    public static string Md(DateOnly? d) => d is { } v ? v.ToString("MM-dd", CultureInfo.InvariantCulture) : "";

    public static string Ymd(DateOnly? d) => d is { } v ? v.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "";

    public static string Ago(DateTime? t)
    {
        if (t is not { } v || v == DateTime.MinValue) return "";
        TimeSpan span = DateTime.Now - v;
        if (span.TotalMinutes < 1) return "just now";
        if (span.TotalMinutes < 60) return $"{(int)span.TotalMinutes}m ago";
        if (span.TotalHours < 24) return $"{(int)span.TotalHours}h ago";
        if (span.TotalDays < 7) return $"{(int)span.TotalDays}d ago";
        return v.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    }

    public static string Stamp(DateTime t) => t.ToString("MM-dd HH:mm", CultureInfo.InvariantCulture);

    public static string Size(long bytes) =>
        bytes >= 1024 * 1024 ? $"{bytes / 1024d / 1024d:0.0} MB" : $"{Math.Max(1, bytes / 1024)} KB";

    public static DateOnly Today => DateOnly.FromDateTime(DateTime.Today);
    /// <summary>
    /// Minimal markdown for AI summaries: headings, bullet/numbered lists, pipe tables, **bold**,
    /// `code` and paragraphs. Everything is HTML-encoded first, so model output cannot inject markup.
    /// </summary>
    public static string MarkdownToHtml(string? markdown)
    {
        var html = new StringBuilder();
        string[] lines = (markdown ?? "").Replace("\r\n", "\n").Split('\n');
        string? list = null;
        var para = new List<string>();

        void FlushPara()
        {
            if (para.Count == 0) return;
            html.Append("<p>").Append(string.Join("<br/>", para.Select(Inline))).Append("</p>");
            para.Clear();
        }
        void CloseList()
        {
            if (list is null) return;
            html.Append($"</{list}>");
            list = null;
        }

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i].TrimEnd();
            string t = line.TrimStart();
            if (t.Length == 0) { FlushPara(); CloseList(); continue; }

            Match h = Regex.Match(t, @"^(#{1,6})\s+(.*)$");
            if (h.Success)
            {
                FlushPara(); CloseList();
                int level = Math.Min(6, h.Groups[1].Length + 2);
                html.Append($"<h{level}>").Append(Inline(h.Groups[2].Value)).Append($"</h{level}>");
                continue;
            }

            if (t.StartsWith('|') && i + 1 < lines.Length && Regex.IsMatch(lines[i + 1].Trim(), @"^\|?\s*:?-{2,}"))
            {
                FlushPara(); CloseList();
                html.Append("<table><thead><tr>");
                foreach (string c in Cells(t)) html.Append("<th>").Append(Inline(c)).Append("</th>");
                html.Append("</tr></thead><tbody>");
                i += 2;
                for (; i < lines.Length && lines[i].TrimStart().StartsWith('|'); i++)
                {
                    html.Append("<tr>");
                    foreach (string c in Cells(lines[i].Trim())) html.Append("<td>").Append(Inline(c)).Append("</td>");
                    html.Append("</tr>");
                }
                i--;
                html.Append("</tbody></table>");
                continue;
            }

            Match b = Regex.Match(t, @"^([-*+]|\d+[.)])\s+(.*)$");
            if (b.Success)
            {
                FlushPara();
                string kind = char.IsDigit(b.Groups[1].Value[0]) ? "ol" : "ul";
                if (list != kind) { CloseList(); html.Append($"<{kind}>"); list = kind; }
                html.Append("<li>").Append(Inline(b.Groups[2].Value)).Append("</li>");
                continue;
            }

            CloseList();
            para.Add(t);
        }
        FlushPara();
        CloseList();
        return html.ToString();

        static IEnumerable<string> Cells(string row) =>
            row.Trim().Trim('|').Split('|').Select(c => c.Trim());

        static string Inline(string s)
        {
            string e = WebUtility.HtmlEncode(s);
            e = Regex.Replace(e, @"`([^`]+)`", "<code>$1</code>");
            e = Regex.Replace(e, @"\*\*([^*]+)\*\*", "<strong>$1</strong>");
            return e;
        }
    }
}
