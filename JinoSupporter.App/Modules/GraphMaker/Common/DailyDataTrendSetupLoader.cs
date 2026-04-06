using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace GraphMaker;

internal static class DailyDataTrendSetupLoader
{
    public static ProcessTrendFileInfo LoadProcessTrendFileInfo(
        IEnumerable<string> filePaths,
        string delimiter,
        int headerRowNumber,
        bool requireFirstColumnDate,
        bool convertWideToSingleY)
    {
        string[] paths = filePaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        if (paths.Length == 0)
        {
            throw new InvalidOperationException("Select at least one file.");
        }

        DataTable mergedTable = LoadSingleTable(paths[0], delimiter, headerRowNumber);
        for (int i = 1; i < paths.Length; i++)
        {
            DataTable extraTable = LoadSingleTable(paths[i], delimiter, headerRowNumber);
            MergeTable(mergedTable, extraTable);
        }

        if (requireFirstColumnDate)
        {
            EnsureFirstColumnIsDate(mergedTable);
        }

        if (convertWideToSingleY)
        {
            mergedTable = ConvertWideToSingleYTable(mergedTable);
        }

        return new ProcessTrendFileInfo
        {
            Name = string.Join(" + ", paths.Select(Path.GetFileName)),
            FilePath = string.Join("|", paths),
            Delimiter = delimiter,
            HeaderRowNumber = headerRowNumber,
            FullData = mergedTable
        };
    }

    private static DataTable LoadSingleTable(string filePath, string delimiter, int headerRowNumber)
    {
        string[] lines = File.ReadAllLines(filePath);
        if (lines.Length == 0)
        {
            throw new InvalidOperationException($"File is empty: {Path.GetFileName(filePath)}");
        }

        int headerIndex = headerRowNumber - 1;
        if (headerIndex < 0 || headerIndex >= lines.Length)
        {
            throw new InvalidOperationException($"Invalid header row number for file: {Path.GetFileName(filePath)}");
        }

        string[] headerTokens = GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter);
        if (headerTokens.Length == 0)
        {
            throw new InvalidOperationException($"Cannot read header from file: {Path.GetFileName(filePath)}");
        }

        DataTable table = new();
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
            if (values.Length == 0)
            {
                continue;
            }

            DataRow row = table.NewRow();
            for (int col = 0; col < Math.Min(values.Length, table.Columns.Count); col++)
            {
                row[col] = values[col];
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private static void MergeTable(DataTable merged, DataTable extra)
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

    private static void EnsureFirstColumnIsDate(DataTable table)
    {
        if (table.Columns.Count == 0)
        {
            throw new InvalidOperationException("Daily Data Trend requires at least one column.");
        }

        string firstColumnName = table.Columns[0].ColumnName;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            string text = table.Rows[rowIndex][0]?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) ||
                DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
            {
                continue;
            }

            throw new InvalidOperationException(
                $"Daily Data Trend requires the first column to contain date values. Invalid value '{text}' was found at row {rowIndex + 1:N0} in column '{firstColumnName}'.");
        }
    }

    private static DataTable ConvertWideToSingleYTable(DataTable wideTable)
    {
        if (wideTable.Columns.Count < 2)
        {
            throw new InvalidOperationException("SingleX(Date) - SingleY requires at least one date column and one value column.");
        }

        DataTable longTable = new();
        longTable.Columns.Add("Date");
        longTable.Columns.Add("Value");

        foreach (DataRow sourceRow in wideTable.Rows)
        {
            string dateText = sourceRow[0]?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(dateText))
            {
                continue;
            }

            for (int columnIndex = 1; columnIndex < wideTable.Columns.Count; columnIndex++)
            {
                string valueText = sourceRow[columnIndex]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(valueText))
                {
                    continue;
                }

                DataRow newRow = longTable.NewRow();
                newRow["Date"] = dateText;
                newRow["Value"] = valueText;
                longTable.Rows.Add(newRow);
            }
        }

        return longTable;
    }
}
