using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace GraphMaker;

public static class GraphMakerFileViewHelper
{
    public static FileInfo_DailySampling CreateFileInfo(string filePath, string delimiter, int headerRowNumber)
    {
        return new FileInfo_DailySampling
        {
            Name = Path.GetFileName(filePath),
            FilePath = filePath,
            Delimiter = delimiter,
            HeaderRowNumber = headerRowNumber
        };
    }

    public static void ParseHeaderIntoDataTable(FileInfo_DailySampling file, int minimumColumnCount)
    {
        file.Delimiter ??= "\t";
        if (file.HeaderRowNumber <= 0)
        {
            file.HeaderRowNumber = 1;
        }

        string[] lines = File.ReadAllLines(file.FilePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException("The selected file is empty.");
        }

        int headerIndex = Math.Max(0, file.HeaderRowNumber - 1);
        if (headerIndex >= lines.Length)
        {
            throw new InvalidOperationException("Header row is out of range.");
        }

        string[] headers = GraphMakerTableHelper.SplitLine(lines[headerIndex], file.Delimiter);
        if (headers.Length < minimumColumnCount)
        {
            throw new InvalidOperationException(minimumColumnCount >= 2
                ? "At least 2 columns are required (X + data column)."
                : "At least 1 data column is required.");
        }

        var uniqueHeaders = GraphMakerTableHelper.BuildUniqueHeaders(headers);
        var table = new DataTable();
        foreach (string header in uniqueHeaders)
        {
            table.Columns.Add(header);
        }

        for (int i = headerIndex + 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            string[] values = GraphMakerTableHelper.SplitLine(lines[i], file.Delimiter);
            DataRow row = table.NewRow();
            for (int c = 0; c < table.Columns.Count; c++)
            {
                row[c] = c < values.Length ? values[c].Trim() : string.Empty;
            }

            table.Rows.Add(row);
        }

        file.FullData = table;
    }

    public static void ParseDateNoHeaderIntoDataTable(FileInfo_DailySampling file)
    {
        file.Delimiter ??= "\t";

        string[] lines = File.ReadAllLines(file.FilePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException("The selected file is empty.");
        }

        var tokenRows = new List<string[]>();
        int maxColumnCount = 0;

        foreach (string line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            string[] tokens = GraphMakerTableHelper.SplitLine(line, file.Delimiter)
                .Select(token => token.Trim())
                .ToArray();

            if (tokens.Length == 0)
            {
                continue;
            }

            string dateText = tokens[0];
            if (!GraphMakerParsingHelper.TryParseDate(dateText, out _))
            {
                throw new InvalidOperationException($"First column must be Date. Invalid value '{dateText}' in file '{Path.GetFileName(file.FilePath)}'.");
            }

            tokenRows.Add(tokens);
            maxColumnCount = Math.Max(maxColumnCount, tokens.Length);
        }

        if (tokenRows.Count == 0 || maxColumnCount < 2)
        {
            throw new InvalidOperationException("At least one date column and one measurement column are required.");
        }

        var table = new DataTable();
        table.Columns.Add("Date");
        for (int i = 1; i < maxColumnCount; i++)
        {
            table.Columns.Add($"Value{i}");
        }

        foreach (string[] tokens in tokenRows)
        {
            DataRow row = table.NewRow();
            for (int i = 0; i < Math.Min(tokens.Length, table.Columns.Count); i++)
            {
                row[i] = tokens[i];
            }

            table.Rows.Add(row);
        }

        file.FullData = table;
    }

    public static void EnsureColumnLimits(FileInfo_DailySampling file, Func<string, bool>? includeColumn = null)
    {
        if (file.FullData == null)
        {
            return;
        }

        includeColumn ??= static _ => true;

        foreach (DataColumn column in file.FullData.Columns)
        {
            if (!includeColumn(column.ColumnName))
            {
                continue;
            }

            if (!file.SavedColumnLimits.ContainsKey(column.ColumnName))
            {
                file.SavedColumnLimits[column.ColumnName] = new ColumnLimitSetting
                {
                    ColumnName = column.ColumnName
                };
            }
        }

        var stale = file.SavedColumnLimits.Keys
            .Where(key => !file.FullData.Columns.Contains(key) || !includeColumn(key))
            .ToList();

        foreach (string key in stale)
        {
            file.SavedColumnLimits.Remove(key);
        }
    }

    public static void BindColumnOptions(
        FileInfo_DailySampling file,
        ObservableCollection<SelectableColumnOption> columnOptions,
        Dictionary<string, HashSet<string>> selectedColumnsByFile,
        Func<string, bool>? includeColumn = null)
    {
        columnOptions.Clear();
        if (file.FullData == null)
        {
            return;
        }

        includeColumn ??= static _ => true;
        string[] availableColumns = file.FullData.Columns.Cast<DataColumn>()
            .Select(column => column.ColumnName)
            .Where(includeColumn)
            .ToArray();

        if (!selectedColumnsByFile.TryGetValue(file.FilePath, out HashSet<string>? selected))
        {
            selected = new HashSet<string>(availableColumns, StringComparer.OrdinalIgnoreCase);
            selectedColumnsByFile[file.FilePath] = selected;
        }

        foreach (string columnName in availableColumns)
        {
            columnOptions.Add(new SelectableColumnOption
            {
                ColumnName = columnName,
                IsSelected = selected.Contains(columnName)
            });
        }
    }

    public static void SaveCurrentSelectionState(
        FileInfo_DailySampling? currentFile,
        IEnumerable<SelectableColumnOption> columnOptions,
        Dictionary<string, HashSet<string>> selectedColumnsByFile)
    {
        if (currentFile == null)
        {
            return;
        }

        selectedColumnsByFile[currentFile.FilePath] = new HashSet<string>(
            columnOptions.Where(option => option.IsSelected).Select(option => option.ColumnName),
            StringComparer.OrdinalIgnoreCase);
    }

    // ─────────────────────────────────────────────────────────
    // DailyDataTrend 전용: 여러 파일을 로드+병합+Date/Value 변환
    // ─────────────────────────────────────────────────────────

    /// <summary>
    /// 파일 목록을 로드, 병합, Wide→Long 변환하여 ProcessTrendFileInfo를 반환합니다.
    /// </summary>
    public static ProcessTrendFileInfo LoadAndMergeDailyFiles(
        IEnumerable<string> filePaths,
        string delimiter,
        int headerRowNumber)
    {
        string[] paths = filePaths
            .Where(p => !string.IsNullOrWhiteSpace(p) && File.Exists(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (paths.Length == 0)
        {
            throw new InvalidOperationException("No valid files found.");
        }

        DataTable merged = LoadDailyTable(paths[0], delimiter, headerRowNumber);
        for (int i = 1; i < paths.Length; i++)
        {
            MergeTables(merged, LoadDailyTable(paths[i], delimiter, headerRowNumber));
        }

        DataTable converted = ConvertToDateValueTable(merged);
        return new ProcessTrendFileInfo
        {
            Name = string.Join(" + ", paths.Select(Path.GetFileName)),
            FilePath = string.Join("|", paths),
            Delimiter = delimiter,
            HeaderRowNumber = headerRowNumber,
            FullData = converted
        };
    }

    private static DataTable LoadDailyTable(string filePath, string delimiter, int headerRowNumber)
    {
        string[] lines = File.ReadAllLines(filePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException($"File is empty: {Path.GetFileName(filePath)}");
        }

        int headerIndex = Math.Max(0, headerRowNumber - 1);
        if (headerIndex >= lines.Length)
        {
            throw new InvalidOperationException($"Header row {headerRowNumber} exceeds file length: {Path.GetFileName(filePath)}");
        }

        string[] headerTokens = GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter);
        var table = new DataTable();
        foreach (string header in GraphMakerTableHelper.BuildUniqueHeaders(headerTokens))
        {
            table.Columns.Add(header);
        }

        for (int i = headerIndex + 1; i < lines.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(lines[i]))
            {
                continue;
            }

            string[] values = GraphMakerTableHelper.SplitLine(lines[i], delimiter);
            DataRow row = table.NewRow();
            for (int c = 0; c < Math.Min(values.Length, table.Columns.Count); c++)
            {
                row[c] = values[c];
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private static void MergeTables(DataTable merged, DataTable extra)
    {
        foreach (DataColumn column in extra.Columns)
        {
            if (!merged.Columns.Contains(column.ColumnName))
            {
                merged.Columns.Add(column.ColumnName);
            }
        }

        foreach (DataRow extraRow in extra.Rows)
        {
            DataRow newRow = merged.NewRow();
            foreach (DataColumn column in extra.Columns)
            {
                newRow[column.ColumnName] = extraRow[column]?.ToString() ?? string.Empty;
            }

            merged.Rows.Add(newRow);
        }
    }

    private static DataTable ConvertToDateValueTable(DataTable wide)
    {
        if (wide.Columns.Count < 2)
        {
            throw new InvalidOperationException("At least one Date column and one value column are required.");
        }

        var result = new DataTable();
        result.Columns.Add("Date");
        result.Columns.Add("Value");

        foreach (DataRow sourceRow in wide.Rows)
        {
            string dateText = sourceRow[0]?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(dateText))
            {
                continue;
            }

            if (!DateTime.TryParse(dateText, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) &&
                !DateTime.TryParse(dateText, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
            {
                throw new InvalidOperationException(
                    $"First column must contain date values. Invalid value: '{dateText}'");
            }

            for (int c = 1; c < wide.Columns.Count; c++)
            {
                string valueText = sourceRow[c]?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(valueText))
                {
                    DataRow row = result.NewRow();
                    row["Date"] = dateText;
                    row["Value"] = valueText;
                    result.Rows.Add(row);
                }
            }
        }

        return result;
    }
}
