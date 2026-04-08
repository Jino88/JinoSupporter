using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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

    /// <summary>
    /// 파일 내용을 분석하여 구분자를 자동 감지합니다 (Tab / Comma / Space).
    /// 컬럼 수가 많고 일관성이 높은 구분자를 반환합니다.
    /// </summary>
    public static string DetectDelimiter(string filePath, int checkLines = 20)
    {
        var sampleLines = File.ReadLines(filePath)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Take(checkLines)
            .ToList();

        if (sampleLines.Count == 0)
        {
            return "\t";
        }

        string[] candidates = { "\t", ",", " " };
        string bestDelimiter = "\t";
        double bestScore = -1;

        foreach (string delimiter in candidates)
        {
            var counts = sampleLines.Select(line => SplitLine(line, delimiter).Length).ToList();
            double mean = counts.Average();
            if (mean < 2)
            {
                continue;
            }

            double variance = counts.Sum(c => Math.Pow(c - mean, 2)) / counts.Count;
            double score = mean / (1.0 + Math.Sqrt(variance));

            if (score > bestScore)
            {
                bestScore = score;
                bestDelimiter = delimiter;
            }
        }

        return bestDelimiter;
    }
}
