using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace GraphMaker
{
    public sealed class NoXMultiYPoint
    {
        public double X { get; set; }
        public double Y { get; set; }
        public int RowIndex { get; set; }
    }

    public sealed class NoXMultiYColumnResult
    {
        public string ColumnName { get; set; } = string.Empty;
        public List<NoXMultiYPoint> Points { get; set; } = new();
        public List<double> Values { get; set; } = new();
        public int TotalCount { get; set; }
        public int NgCount { get; set; }
        public double NgRatePercent { get; set; }
        public double Avg { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double StdDev { get; set; }
        public double? Cpk { get; set; }
        public double? Spec { get; set; }
        public double? Upper { get; set; }
        public double? Lower { get; set; }
    }

    public sealed class NoXMultiYGraphResult
    {
        public List<string> Categories { get; set; } = new();
        public List<NoXMultiYColumnResult> Columns { get; set; } = new();
    }

    public static class NoXMultiYGraphCalculator
    {
        public static NoXMultiYGraphResult Calculate(FileInfo_DailySampling file, List<string> selectedColumns)
        {
            if (file.FullData == null)
            {
                throw new InvalidOperationException("No data loaded.");
            }

            var table = file.FullData;
            var result = new NoXMultiYGraphResult
            {
                Categories = selectedColumns
                    .Where(columnName => table.Columns.Contains(columnName))
                    .ToList()
            };

            for (int columnOrder = 0; columnOrder < result.Categories.Count; columnOrder++)
            {
                string columnName = result.Categories[columnOrder];
                int columnIndex = table.Columns[columnName].Ordinal;
                var limits = ReadLimits(file, columnName);
                var values = new List<double>();
                var points = new List<NoXMultiYPoint>();
                int ngCount = 0;

                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    string text = table.Rows[rowIndex][columnIndex]?.ToString() ?? string.Empty;
                    if (!GraphMakerParsingHelper.TryParseDouble(text, out double y))
                    {
                        continue;
                    }

                    values.Add(y);
                    points.Add(new NoXMultiYPoint
                    {
                        X = columnOrder,
                        Y = y,
                        RowIndex = rowIndex
                    });

                    if (limits.upper.HasValue && limits.lower.HasValue && (y > limits.upper.Value || y < limits.lower.Value))
                    {
                        ngCount++;
                    }
                }

                double avg = values.Count > 0 ? values.Average() : 0.0;
                double variance = values.Count > 0
                    ? values.Sum(v => Math.Pow(v - avg, 2)) / values.Count
                    : 0.0;
                double stdDev = Math.Sqrt(variance);
                double? cpk = null;
                if (values.Count > 0 && stdDev > 0 && limits.upper.HasValue && limits.lower.HasValue)
                {
                    double cpu = (limits.upper.Value - avg) / (3.0 * stdDev);
                    double cpl = (avg - limits.lower.Value) / (3.0 * stdDev);
                    cpk = Math.Min(cpu, cpl);
                }

                result.Columns.Add(new NoXMultiYColumnResult
                {
                    ColumnName = columnName,
                    Points = points,
                    Values = values,
                    TotalCount = values.Count,
                    NgCount = ngCount,
                    NgRatePercent = values.Count > 0 ? (double)ngCount / values.Count * 100.0 : 0.0,
                    Avg = avg,
                    Min = values.Count > 0 ? values.Min() : 0.0,
                    Max = values.Count > 0 ? values.Max() : 0.0,
                    StdDev = stdDev,
                    Cpk = cpk,
                    Spec = limits.spec,
                    Upper = limits.upper,
                    Lower = limits.lower
                });
            }

            return result;
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
    }
}
