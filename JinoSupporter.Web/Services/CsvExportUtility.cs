using System.Text;

namespace JinoSupporter.Web.Services;

public static class CsvExportUtility
{
    public const string ContentType = "text/csv;charset=utf-8";

    public static byte[] Build(IEnumerable<IEnumerable<string?>> rows)
    {
        var sb = new StringBuilder();
        foreach (var row in rows)
            AppendRow(sb, row);

        return ToUtf8BomBytes(sb);
    }

    public static byte[] ToUtf8BomBytes(StringBuilder sb)
        => Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes(sb.ToString()))
            .ToArray();

    public static void AppendRow(StringBuilder sb, IEnumerable<string?> values)
    {
        bool first = true;
        foreach (string? value in values)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(Escape(value));
        }

        sb.AppendLine();
    }

    public static string Escape(string? value)
    {
        string text = value ?? string.Empty;
        if (!text.Contains('"') && !text.Contains(',') && !text.Contains('\r') && !text.Contains('\n'))
            return text;

        return "\"" + text.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
    }
}
