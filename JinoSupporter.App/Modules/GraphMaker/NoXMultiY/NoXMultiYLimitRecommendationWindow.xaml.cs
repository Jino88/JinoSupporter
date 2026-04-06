using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;

namespace GraphMaker
{
    public partial class NoXMultiYLimitRecommendationWindow : Window
    {
        private sealed class RecommendationRow
        {
            public string ColumnName { get; init; } = string.Empty;
            public int Count { get; init; }
            public double Avg { get; init; }
            public double StdDev { get; init; }
            public double CurrentNgRatePercent { get; init; }
            public double RequiredLowerActual { get; init; }
            public double RequiredUpperActual { get; init; }
            public double TrendAvg { get; init; }
            public double TrendStdDev { get; init; }
            public double RequiredLowerTrend { get; init; }
            public double RequiredUpperTrend { get; init; }
        }

        private readonly NoXMultiYGraphResult _result;

        public NoXMultiYLimitRecommendationWindow(string fileName, NoXMultiYGraphResult result)
        {
            InitializeComponent();
            _result = result;
            HeaderTextBlock.Text = $"{fileName} - Recommended USL / LSL for all areas";
            TargetNgRateTextBox.Text = "1.0";
            Refresh();
        }

        private void TargetNgRateTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            Refresh();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Refresh()
        {
            double targetNgRatePercent = ParseTargetNgRate();
            double targetYield = Math.Clamp(1.0 - (targetNgRatePercent / 100.0), 0.000001, 0.999999);
            double zValue = InverseStandardNormalCdf((1.0 + targetYield) / 2.0);

            var rows = BuildRows(zValue);
            RecommendationGrid.ItemsSource = rows;

            var actualLower = rows.Min(r => r.RequiredLowerActual);
            var actualUpper = rows.Max(r => r.RequiredUpperActual);
            var trendLower = rows.Min(r => r.RequiredLowerTrend);
            var trendUpper = rows.Max(r => r.RequiredUpperTrend);

            ActualSummaryTextBlock.Text = BuildActualSummary(targetNgRatePercent, zValue, actualLower, actualUpper, rows);
            TrendSummaryTextBlock.Text = BuildTrendSummary(targetNgRatePercent, zValue, trendLower, trendUpper, rows);
        }

        private List<RecommendationRow> BuildRows(double zValue)
        {
            var avgRegression = BuildLinearRegression(_result.Columns.Select((column, index) => (x: (double)index, y: column.Avg)));
            var stdRegression = BuildLinearRegression(_result.Columns.Select((column, index) => (x: (double)index, y: Math.Max(column.StdDev, 0.0000001))));

            var rows = new List<RecommendationRow>();
            for (int index = 0; index < _result.Columns.Count; index++)
            {
                var column = _result.Columns[index];
                double requiredLowerActual = column.Avg - (zValue * Math.Max(column.StdDev, 0.0000001));
                double requiredUpperActual = column.Avg + (zValue * Math.Max(column.StdDev, 0.0000001));
                double trendAvg = avgRegression.Intercept + (avgRegression.Slope * index);
                double trendStdDev = Math.Max(stdRegression.Intercept + (stdRegression.Slope * index), 0.0000001);
                double requiredLowerTrend = trendAvg - (zValue * trendStdDev);
                double requiredUpperTrend = trendAvg + (zValue * trendStdDev);

                rows.Add(new RecommendationRow
                {
                    ColumnName = column.ColumnName,
                    Count = column.TotalCount,
                    Avg = column.Avg,
                    StdDev = column.StdDev,
                    CurrentNgRatePercent = column.NgRatePercent,
                    RequiredLowerActual = requiredLowerActual,
                    RequiredUpperActual = requiredUpperActual,
                    TrendAvg = trendAvg,
                    TrendStdDev = trendStdDev,
                    RequiredLowerTrend = requiredLowerTrend,
                    RequiredUpperTrend = requiredUpperTrend
                });
            }

            return rows;
        }

        private double ParseTargetNgRate()
        {
            if (!GraphMakerParsingHelper.TryParseDouble(TargetNgRateTextBox.Text, out double targetNgRatePercent))
            {
                targetNgRatePercent = 1.0;
            }

            return Math.Clamp(targetNgRatePercent, 0.0001, 99.0);
        }

        private static string BuildActualSummary(double targetNgRatePercent, double zValue, double lower, double upper, List<RecommendationRow> rows)
        {
            string widestArea = rows
                .OrderByDescending(r => r.RequiredUpperActual - r.RequiredLowerActual)
                .FirstOrDefault()?.ColumnName ?? "-";

            var sb = new StringBuilder();
            sb.AppendLine($"Target NG rate: {targetNgRatePercent:F4}%");
            sb.AppendLine($"Applied z-value: {zValue:F4}");
            sb.AppendLine($"Recommended common LSL: {lower:F4}");
            sb.AppendLine($"Recommended common USL: {upper:F4}");
            sb.AppendLine($"Widest actual envelope: {widestArea}");
            sb.Append("Rule: each area uses Avg ± z * StdDev, then the common band takes the outermost LSL/USL.");
            return sb.ToString();
        }

        private static string BuildTrendSummary(double targetNgRatePercent, double zValue, double lower, double upper, List<RecommendationRow> rows)
        {
            double avgSlope = BuildLinearRegression(rows.Select((row, index) => (x: (double)index, y: row.Avg))).Slope;
            double stdSlope = BuildLinearRegression(rows.Select((row, index) => (x: (double)index, y: Math.Max(row.StdDev, 0.0000001)))).Slope;

            var sb = new StringBuilder();
            sb.AppendLine($"Target NG rate: {targetNgRatePercent:F4}%");
            sb.AppendLine($"Applied z-value: {zValue:F4}");
            sb.AppendLine($"Trend-based common LSL: {lower:F4}");
            sb.AppendLine($"Trend-based common USL: {upper:F4}");
            sb.AppendLine($"Avg slope per area: {avgSlope:F6}");
            sb.AppendLine($"StdDev slope per area: {stdSlope:F6}");
            sb.Append("Rule: linear regression is fitted on area index for Avg and StdDev, then predicted envelopes are merged.");
            return sb.ToString();
        }

        private static (double Slope, double Intercept) BuildLinearRegression(IEnumerable<(double x, double y)> points)
        {
            var list = points.ToList();
            if (list.Count == 0)
            {
                return (0.0, 0.0);
            }

            if (list.Count == 1)
            {
                return (0.0, list[0].y);
            }

            double avgX = list.Average(p => p.x);
            double avgY = list.Average(p => p.y);
            double numerator = list.Sum(p => (p.x - avgX) * (p.y - avgY));
            double denominator = list.Sum(p => Math.Pow(p.x - avgX, 2));
            double slope = Math.Abs(denominator) < 0.0000001 ? 0.0 : numerator / denominator;
            double intercept = avgY - (slope * avgX);
            return (slope, intercept);
        }

        private static double InverseStandardNormalCdf(double probability)
        {
            probability = Math.Clamp(probability, 0.0000001, 0.9999999);

            double[] a =
            {
                -3.969683028665376e+01,
                2.209460984245205e+02,
                -2.759285104469687e+02,
                1.383577518672690e+02,
                -3.066479806614716e+01,
                2.506628277459239e+00
            };
            double[] b =
            {
                -5.447609879822406e+01,
                1.615858368580409e+02,
                -1.556989798598866e+02,
                6.680131188771972e+01,
                -1.328068155288572e+01
            };
            double[] c =
            {
                -7.784894002430293e-03,
                -3.223964580411365e-01,
                -2.400758277161838e+00,
                -2.549732539343734e+00,
                4.374664141464968e+00,
                2.938163982698783e+00
            };
            double[] d =
            {
                7.784695709041462e-03,
                3.224671290700398e-01,
                2.445134137142996e+00,
                3.754408661907416e+00
            };

            const double pLow = 0.02425;
            const double pHigh = 1.0 - pLow;

            if (probability < pLow)
            {
                double q = Math.Sqrt(-2.0 * Math.Log(probability));
                return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                       ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0);
            }

            if (probability <= pHigh)
            {
                double q = probability - 0.5;
                double r = q * q;
                return (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q /
                       (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0);
            }

            double tailQ = Math.Sqrt(-2.0 * Math.Log(1.0 - probability));
            return -(((((c[0] * tailQ + c[1]) * tailQ + c[2]) * tailQ + c[3]) * tailQ + c[4]) * tailQ + c[5]) /
                    ((((d[0] * tailQ + d[1]) * tailQ + d[2]) * tailQ + d[3]) * tailQ + 1.0);
        }
    }
}
