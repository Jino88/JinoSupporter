using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;

namespace GraphMaker
{
    public enum XAxisValueType
    {
        Numeric,
        Date,
        Category
    }

    public sealed class ColumnGraphRow
    {
        public int RowIndex { get; set; }
        public double Y { get; set; }
    }

    public sealed class ColumnGraphResult
    {
        public string ColumnName { get; set; } = string.Empty;
        public List<ColumnGraphRow> Rows { get; set; } = new();

        public double? Spec { get; set; }
        public double? Upper { get; set; }
        public double? Lower { get; set; }

        public int TotalCount { get; set; }
        public int NgCount { get; set; }
        public double NgRatePercent { get; set; }
        public double Avg { get; set; }
        public double StdDev { get; set; }
        public double? Cpk { get; set; }
    }

    public sealed class OverallGraphResult
    {
        public int TotalRows { get; set; }
        public int NgRows { get; set; }
        public double NgRatePercent { get; set; }
    }

    public sealed class MultiColumnGraphResult
    {
        public XAxisValueType XAxisType { get; set; }
        public List<double> XNumericValues { get; set; } = new();
        public List<DateTime> XDateValues { get; set; } = new();
        public List<string> XLabels { get; set; } = new();

        public List<ColumnGraphResult> Columns { get; set; } = new();
        public OverallGraphResult Overall { get; set; } = new();
    }

    public static class MultiColumnGraphCalculator
    {
        public static MultiColumnGraphResult Calculate(FileInfo_DailySampling file, List<string> selectedColumns)
        {
            if (file.FullData == null)
            {
                throw new InvalidOperationException("No data loaded.");
            }

            var table = file.FullData;
            if (table.Columns.Count < 2)
            {
                throw new InvalidOperationException("At least 2 columns are required.");
            }

            int rowCount = table.Rows.Count;
            var result = new MultiColumnGraphResult
            {
                XAxisType = DetectXAxisType(file, table, rowCount)
            };

            BuildXAxisValues(table, rowCount, result);

            var columnIndexMap = selectedColumns
                .Where(c => table.Columns.Contains(c))
                .ToDictionary(c => c, c => table.Columns[c].Ordinal, StringComparer.OrdinalIgnoreCase);

            foreach (var kv in columnIndexMap)
            {
                string columnName = kv.Key;
                int colIndex = kv.Value;
                var limits = ReadLimits(file, columnName);

                var values = new List<double>();
                var rows = new List<ColumnGraphRow>();
                int ngCount = 0;

                for (int r = 0; r < rowCount; r++)
                {
                    string text = table.Rows[r][colIndex]?.ToString() ?? string.Empty;
                    if (!GraphMakerParsingHelper.TryParseDouble(text, out double y))
                    {
                        continue;
                    }

                    values.Add(y);
                    rows.Add(new ColumnGraphRow { RowIndex = r, Y = y });

                    if (limits.upper.HasValue && limits.lower.HasValue && (y > limits.upper.Value || y < limits.lower.Value))
                    {
                        ngCount++;
                    }
                }

                double avg = values.Count > 0 ? values.Average() : 0.0;
                double stdDev = 0.0;
                if (values.Count > 0)
                {
                    double variance = values.Sum(v => Math.Pow(v - avg, 2)) / values.Count;
                    stdDev = Math.Sqrt(variance);
                }

                double? cpk = null;
                if (values.Count > 0 && stdDev > 0 && limits.upper.HasValue && limits.lower.HasValue)
                {
                    double cpu = (limits.upper.Value - avg) / (3.0 * stdDev);
                    double cpl = (avg - limits.lower.Value) / (3.0 * stdDev);
                    cpk = Math.Min(cpu, cpl);
                }

                result.Columns.Add(new ColumnGraphResult
                {
                    ColumnName = columnName,
                    Rows = rows,
                    Spec = limits.spec,
                    Upper = limits.upper,
                    Lower = limits.lower,
                    TotalCount = values.Count,
                    NgCount = ngCount,
                    NgRatePercent = values.Count > 0 ? (double)ngCount / values.Count * 100.0 : 0.0,
                    Avg = avg,
                    StdDev = stdDev,
                    Cpk = cpk
                });
            }

            result.Overall = CalculateOverall(table, file, result.Columns);
            return result;
        }

        private static OverallGraphResult CalculateOverall(DataTable table, FileInfo_DailySampling file, List<ColumnGraphResult> columns)
        {
            int totalRows = table.Rows.Count;
            int ngRows = 0;
            var selectedColumnNames = new HashSet<string>(columns.Select(c => c.ColumnName), StringComparer.OrdinalIgnoreCase);

            for (int r = 0; r < totalRows; r++)
            {
                bool rowNg = false;

                foreach (DataColumn col in table.Columns)
                {
                    if (col.Ordinal == 0 || !selectedColumnNames.Contains(col.ColumnName))
                    {
                        continue;
                    }

                    var limits = ReadLimits(file, col.ColumnName);
                    if (!limits.upper.HasValue || !limits.lower.HasValue)
                    {
                        continue;
                    }

                    string valueText = table.Rows[r][col.Ordinal]?.ToString() ?? string.Empty;
                    if (!GraphMakerParsingHelper.TryParseDouble(valueText, out double value))
                    {
                        continue;
                    }

                    if (value > limits.upper.Value || value < limits.lower.Value)
                    {
                        rowNg = true;
                        break;
                    }
                }

                if (rowNg)
                {
                    ngRows++;
                }
            }

            return new OverallGraphResult
            {
                TotalRows = totalRows,
                NgRows = ngRows,
                NgRatePercent = totalRows > 0 ? (double)ngRows / totalRows * 100.0 : 0.0
            };
        }

        private static (double? spec, double? upper, double? lower) ReadLimits(FileInfo_DailySampling file, string columnName)
        {
            if (!file.SavedColumnLimits.TryGetValue(columnName, out var setting))
            {
                return (null, null, null);
            }

            double? spec = GraphMakerParsingHelper.TryParseDouble(setting.SpecValue, out double s) ? s : null;
            double? upper = GraphMakerParsingHelper.TryParseDouble(setting.UpperValue, out double u) ? u : null;
            double? lower = GraphMakerParsingHelper.TryParseDouble(setting.LowerValue, out double l) ? l : null;
            return (spec, upper, lower);
        }

        private static XAxisValueType DetectXAxisType(FileInfo_DailySampling file, DataTable table, int rowCount)
        {
            if (file.ForcedXAxisType.HasValue)
            {
                return file.ForcedXAxisType.Value;
            }

            if (rowCount == 0)
            {
                return XAxisValueType.Category;
            }

            int numericCount = 0;
            int dateCount = 0;
            int validCount = 0;

            for (int r = 0; r < rowCount; r++)
            {
                string text = table.Rows[r][0]?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                validCount++;
                if (GraphMakerParsingHelper.TryParseDouble(text, out _))
                {
                    numericCount++;
                }

                if (GraphMakerParsingHelper.TryParseDate(text, out _))
                {
                    dateCount++;
                }
            }

            if (validCount > 0 && numericCount == validCount)
            {
                return XAxisValueType.Numeric;
            }

            if (validCount > 0 && dateCount == validCount)
            {
                return XAxisValueType.Date;
            }

            return XAxisValueType.Category;
        }

        private static void BuildXAxisValues(DataTable table, int rowCount, MultiColumnGraphResult result)
        {
            for (int r = 0; r < rowCount; r++)
            {
                string text = table.Rows[r][0]?.ToString() ?? string.Empty;
                switch (result.XAxisType)
                {
                    case XAxisValueType.Numeric:
                        result.XNumericValues.Add(GraphMakerParsingHelper.TryParseDouble(text, out double n) ? n : r + 1);
                        result.XLabels.Add(text);
                        break;
                    case XAxisValueType.Date:
                        result.XDateValues.Add(GraphMakerParsingHelper.TryParseDate(text, out DateTime d) ? d : DateTime.MinValue.AddDays(r));
                        result.XLabels.Add(text);
                        break;
                    default:
                        result.XLabels.Add(string.IsNullOrWhiteSpace(text) ? (r + 1).ToString(CultureInfo.InvariantCulture) : text);
                        break;
                }
            }
        }

    }
}
