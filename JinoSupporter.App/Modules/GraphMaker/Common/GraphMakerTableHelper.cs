using System;
using System.Collections.Generic;

namespace GraphMaker;

public static class GraphMakerTableHelper
{
    public static List<string> BuildUniqueHeaders(IReadOnlyList<string> headers)
    {
        var result = new List<string>();
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (string raw in headers)
        {
            string header = string.IsNullOrWhiteSpace(raw) ? "Column" : raw.Trim();

            if (!counts.TryGetValue(header, out int count))
            {
                counts[header] = 1;
                result.Add(header);
                continue;
            }

            count++;
            counts[header] = count;
            result.Add($"{header}_{count}");
        }

        return result;
    }

    public static string[] SplitLine(string line, string delimiter)
    {
        if (delimiter == " ")
        {
            return line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        }

        return line.Split(new[] { delimiter }, StringSplitOptions.None);
    }
}
