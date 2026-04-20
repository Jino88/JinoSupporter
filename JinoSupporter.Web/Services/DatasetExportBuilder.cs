using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace JinoSupporter.Web.Services;

/// <summary>
/// Builds a ZIP bundle of datasets for handing back to Claude (or another reviewer).
/// All files are FLAT — no subfolders — so the bundle reads like a single working
/// directory. Bulk export prefixes each file with the dataset name.
/// </summary>
public static class DatasetExportBuilder
{
    /// <summary>
    /// Single dataset → flat ZIP. No subfolders. Contents:
    ///   _meta.md, _measurements.tsv, _summary.md, _issues.txt,
    ///   img-01.png, img-02.png, …,
    ///   file-XXX.ext (backup files if any)
    /// </summary>
    public static byte[] BuildZip(WebRepository repo, string datasetName)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteDatasetEntries(archive, repo, datasetName, fileNamePrefix: "");
        }
        return ms.ToArray();
    }

    /// <summary>
    /// All datasets with validation issues → one flat ZIP, no per-dataset folders.
    /// Every file is prefixed with <c>{NN}_{sanitizedName}_…</c> so files from the
    /// same dataset sort together.
    /// </summary>
    public static byte[] BuildAllFlaggedZip(WebRepository repo)
    {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var reports = repo.GetAllRawReports();

            // Collect flagged datasets first so we can number them stably.
            var flagged = new List<(RawReportInfo r, List<MeasurementValidator.Issue> issues)>();
            foreach (var r in reports)
            {
                if (r.MeasurementCount == 0) continue;
                var rows   = repo.GetNormalizedMeasurements(r.DatasetName);
                var issues = MeasurementValidator.FindIssues(rows);
                if (issues.Count > 0) flagged.Add((r, issues));
            }

            // INDEX.md — high-level map of the bundle
            var index = new StringBuilder();
            index.AppendLine("# Flagged Datasets — Bulk Export\n");
            index.AppendLine($"- Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            index.AppendLine($"- Total flagged: {flagged.Count}");
            index.AppendLine();
            index.AppendLine("All files are flat (no subfolders). Each dataset's files share a `NN_name_*` prefix.\n");

            for (int i = 0; i < flagged.Count; i++)
            {
                var (r, issues) = flagged[i];
                string idx    = (i + 1).ToString("D2");
                string safe   = Sanitize(r.DatasetName, maxLen: 50);
                string prefix = $"{idx}_{safe}_";

                WriteDatasetEntries(archive, repo, r.DatasetName, prefix);

                index.AppendLine($"## {idx}. {r.DatasetName}");
                index.AppendLine($"- Product: `{r.ProductType}` · Date: `{r.ReportDate}` · Rows: {r.MeasurementCount} · Issues: {issues.Count}");
                index.AppendLine($"- Files prefix: `{prefix}`");
                foreach (var iss in issues)
                    index.AppendLine($"  - [{iss.Kind}] {iss.Variable} — {iss.Message}");
                index.AppendLine();
            }

            AddText(archive, "00_INDEX.md", index.ToString());
        }
        return ms.ToArray();
    }

    // ── Core: emit one dataset's files with a given prefix ────────────────────

    private static void WriteDatasetEntries(ZipArchive archive, WebRepository repo,
                                            string datasetName, string fileNamePrefix)
    {
        var reports      = repo.GetAllRawReports();
        var rep          = reports.FirstOrDefault(r => r.DatasetName == datasetName);
        var measurements = repo.GetNormalizedMeasurements(datasetName);
        var summary      = repo.GetDatasetSummaryRecord(datasetName);
        var images       = repo.GetRawReportImages(datasetName);
        var files        = repo.GetRawReportFileInfos(datasetName);
        var issues       = MeasurementValidator.FindIssues(measurements);

        // ─ meta.md
        var meta = new StringBuilder();
        meta.AppendLine($"# {datasetName}\n");
        if (rep is not null)
        {
            meta.AppendLine($"- Product Type: `{rep.ProductType}`");
            meta.AppendLine($"- Report Date:  `{rep.ReportDate}`");
            meta.AppendLine($"- Images:       {rep.ImageCount}");
            meta.AppendLine($"- Measurements: {rep.MeasurementCount}");
            meta.AppendLine($"- Created:      {rep.CreatedAt}");
        }

        meta.AppendLine("\n## Validation Issues");
        if (issues.Count == 0)
        {
            meta.AppendLine("(none)");
        }
        else
        {
            foreach (var iss in issues)
                meta.AppendLine($"- [{iss.Kind}] {iss.Variable}" +
                    (string.IsNullOrEmpty(iss.VariableDetail) ? "" : $" | {iss.VariableDetail}") +
                    $" — {iss.Message}");
        }

        meta.AppendLine("\n## Files (flat, in this ZIP)");
        meta.AppendLine($"- `{fileNamePrefix}meta.md`           — this file");
        meta.AppendLine($"- `{fileNamePrefix}measurements.tsv`  — AI-extracted rows");
        meta.AppendLine($"- `{fileNamePrefix}summary.md`        — AI summary/findings/tags");
        meta.AppendLine($"- `{fileNamePrefix}issues.txt`        — validator issues (plain text)");
        if (images.Count > 0)
            meta.AppendLine($"- `{fileNamePrefix}img-NN.ext`        — source images ({images.Count})");
        if (files.Count > 0)
            meta.AppendLine($"- `{fileNamePrefix}file-*`            — backup files ({files.Count})");

        AddText(archive, $"{fileNamePrefix}meta.md", meta.ToString());

        // ─ measurements.tsv
        var tsv = new StringBuilder();
        tsv.AppendLine("Variable\tGroup\tLine\tCheckType\tInput\tOK\tNG\tNG%\tDefectType\tCategory\tCount\tIntervention\tVariableDetail");
        foreach (var m in measurements)
        {
            tsv.Append(Esc(m.Variable)).Append('\t')
               .Append(Esc(m.VariableGroup)).Append('\t')
               .Append(Esc(m.Line)).Append('\t')
               .Append(Esc(m.CheckType)).Append('\t')
               .Append(m.InputQty).Append('\t')
               .Append(m.OkQty).Append('\t')
               .Append(m.NgTotal).Append('\t')
               .Append(m.NgRate.ToString("F1")).Append("%\t")
               .Append(Esc(m.DefectType)).Append('\t')
               .Append(Esc(m.DefectCategory)).Append('\t')
               .Append(m.DefectCount).Append('\t')
               .Append(Esc(m.Intervention)).Append('\t')
               .Append(Esc(m.VariableDetail))
               .AppendLine();
        }
        AddText(archive, $"{fileNamePrefix}measurements.tsv", tsv.ToString());

        // ─ summary.md
        var sum = new StringBuilder();
        if (summary is null)
        {
            sum.AppendLine("(no summary record)");
        }
        else
        {
            if (!string.IsNullOrWhiteSpace(summary.Purpose))
            {
                sum.AppendLine("## Purpose");
                sum.AppendLine(summary.Purpose);
                sum.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(summary.TestConditions))
            {
                sum.AppendLine("## Test Conditions");
                sum.AppendLine(summary.TestConditions);
                sum.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(summary.RootCause))
            {
                sum.AppendLine("## Root Cause");
                sum.AppendLine(summary.RootCause);
                sum.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(summary.Decision))
            {
                sum.AppendLine("## Decision");
                sum.AppendLine(summary.Decision);
                sum.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(summary.RecommendedAction))
            {
                sum.AppendLine("## Recommended Action");
                sum.AppendLine(summary.RecommendedAction);
                sum.AppendLine();
            }
            sum.AppendLine("## Summary");
            sum.AppendLine(summary.Summary);
            sum.AppendLine("\n## Key Findings");
            sum.AppendLine(summary.KeyFindings);
            if (summary.Tags.Count > 0)
            {
                sum.AppendLine("\n## Tags");
                foreach (var t in summary.Tags) sum.AppendLine($"- {t}");
            }
        }
        AddText(archive, $"{fileNamePrefix}summary.md", sum.ToString());

        // ─ issues.txt
        var iss_ = new StringBuilder();
        if (issues.Count == 0)
        {
            iss_.AppendLine("no issues");
        }
        else
        {
            foreach (var iss in issues)
                iss_.AppendLine($"[{iss.Kind}] {iss.Variable}" +
                    (string.IsNullOrEmpty(iss.VariableDetail) ? "" : $" | {iss.VariableDetail}") +
                    $" -- {iss.Message}");
        }
        AddText(archive, $"{fileNamePrefix}issues.txt", iss_.ToString());

        // ─ images (flat, numbered)
        for (int i = 0; i < images.Count; i++)
        {
            var img = images[i];
            string ext = img.MediaType switch
            {
                "image/png"  => "png",
                "image/jpeg" => "jpg",
                "image/gif"  => "gif",
                "image/webp" => "webp",
                _            => "bin",
            };
            AddBinary(archive, $"{fileNamePrefix}img-{i + 1:D2}.{ext}", img.Data);
        }

        // ─ backup files (flat)
        foreach (var f in files)
        {
            var bytes = repo.GetRawReportFile(f.Id);
            if (bytes is null) continue;
            string safeName = Sanitize(bytes.Value.FileName, maxLen: 80);
            AddBinary(archive, $"{fileNamePrefix}file-{safeName}", bytes.Value.Data);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static void AddText(ZipArchive archive, string entryName, string content)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        // UTF-8 WITHOUT BOM — some viewers render the BOM as literal "ï»¿" characters.
        using var w = new StreamWriter(entry.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        w.Write(content);
    }

    private static void AddBinary(ZipArchive archive, string entryName, byte[] data)
    {
        var entry = archive.CreateEntry(entryName, CompressionLevel.Fastest);
        using var s = entry.Open();
        s.Write(data, 0, data.Length);
    }

    private static string Esc(string s) =>
        (s ?? "").Replace("\t", " ").Replace("\r", "").Replace("\n", " ");

    private static string Sanitize(string name, int maxLen)
    {
        string safe = Regex.Replace(name ?? "", @"[^\w\-.]+", "_").Trim('_');
        if (safe.Length > maxLen) safe = safe[..maxLen].TrimEnd('_');
        return string.IsNullOrEmpty(safe) ? "unnamed" : safe;
    }
}
