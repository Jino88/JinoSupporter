using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using OxyPlot;

namespace GraphMaker
{
    public partial class MultiColumnLargeDetailWindow : Window
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

        private readonly string _fileName;
        private readonly ColumnGraphResult _column;
        private readonly MultiColumnGraphResult _result;

        public MultiColumnLargeDetailWindow(string fileName, ColumnGraphResult column, MultiColumnGraphResult result, PlotModel model, string statsText)
        {
            InitializeComponent();
            _fileName = fileName;
            _column = column;
            _result = result;

            HeaderTextBlock.Text = $"{fileName} - {column.ColumnName}";
            PlotViewControl.Model = model;
            InfoTextBox.Text = statsText;
            BuildOverallPanel();
        }

        private void BuildOverallPanel()
        {
            OverallTotalRowsTextBox.Text = _result.Overall.TotalRows.ToString("N0", CultureInfo.InvariantCulture);
            OverallNgRowsTextBox.Text = _result.Overall.NgRows.ToString("N0", CultureInfo.InvariantCulture);
            OverallNgRateTextBox.Text = _result.Overall.NgRatePercent.ToString("F2", CultureInfo.InvariantCulture);
            OverallPpmTextBox.Text = (_result.Overall.NgRatePercent * 10000.0).ToString("F2", CultureInfo.InvariantCulture);
        }

        private void RecommendButton_Click(object sender, RoutedEventArgs e)
        {
            var rows = BuildRecommendationRows();
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "No valid data is available for recommendation.", "Recommend USL / LSL",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var targetWindow = new Window
            {
                Title = $"{_column.ColumnName} - Recommended USL / LSL",
                Width = 980,
                Height = 700,
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var targetNgRateTextBox = new System.Windows.Controls.TextBox
            {
                Text = "1.0",
                Width = 120,
                Margin = new Thickness(8, 0, 0, 0)
            };
            var summaryTextBlock = new System.Windows.Controls.TextBlock
            {
                Margin = new Thickness(0, 12, 0, 12),
                TextWrapping = TextWrapping.Wrap
            };
            var grid = new System.Windows.Controls.DataGrid
            {
                AutoGenerateColumns = true,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                IsReadOnly = true,
                ItemsSource = rows
            };

            void Refresh()
            {
                double targetNgRatePercent = ParseTargetNgRate(targetNgRateTextBox.Text);
                double yield = Math.Clamp(1.0 - (targetNgRatePercent / 100.0), 0.000001, 0.999999);
                double zValue = InverseStandardNormalCdf((1.0 + yield) / 2.0);
                var refreshedRows = BuildRecommendationRows(zValue);
                grid.ItemsSource = refreshedRows;

                double actualLower = refreshedRows.Min(r => r.RequiredLowerActual);
                double actualUpper = refreshedRows.Max(r => r.RequiredUpperActual);
                double trendLower = refreshedRows.Min(r => r.RequiredLowerTrend);
                double trendUpper = refreshedRows.Max(r => r.RequiredUpperTrend);

                var sb = new StringBuilder();
                sb.AppendLine($"Target NG rate: {targetNgRatePercent:F4}%");
                sb.AppendLine($"Applied z-value: {zValue:F4}");
                sb.AppendLine($"Actual-based common LSL / USL: {actualLower:F4} / {actualUpper:F4}");
                sb.AppendLine($"Trend-based common LSL / USL: {trendLower:F4} / {trendUpper:F4}");
                sb.Append("Actual uses Avg +/- z * StdDev. Trend uses linear regression on row groups.");
                summaryTextBlock.Text = sb.ToString();
            }

            targetNgRateTextBox.TextChanged += (_, _) => Refresh();

            var topPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal
            };
            topPanel.Children.Add(new System.Windows.Controls.TextBlock
            {
                Text = "Target NG Rate (%)",
                VerticalAlignment = System.Windows.VerticalAlignment.Center
            });
            topPanel.Children.Add(targetNgRateTextBox);

            var root = new System.Windows.Controls.DockPanel { Margin = new Thickness(12) };
            System.Windows.Controls.DockPanel.SetDock(topPanel, System.Windows.Controls.Dock.Top);
            System.Windows.Controls.DockPanel.SetDock(summaryTextBlock, System.Windows.Controls.Dock.Top);
            root.Children.Add(topPanel);
            root.Children.Add(summaryTextBlock);
            root.Children.Add(grid);

            targetWindow.Content = root;
            Refresh();
            targetWindow.ShowDialog();
        }

        private List<RecommendationRow> BuildRecommendationRows(double? zValueOverride = null)
        {
            var rows = _column.Rows
                .GroupBy(r => new
                {
                    Key = GetXAxisValue(r.RowIndex),
                    Label = BuildRowLabel(r.RowIndex)
                })
                .OrderBy(g => g.Key.Key)
                .Select((g, idx) =>
                {
                    double avg = g.Average(x => x.Y);
                    var values = g.Select(x => x.Y).ToList();
                    double variance = values.Count > 0 ? values.Sum(v => Math.Pow(v - avg, 2)) / values.Count : 0.0;
                    return new
                    {
                        Name = g.Key.Label,
                        Index = idx,
                        Count = values.Count,
                        Avg = avg,
                        StdDev = Math.Sqrt(variance)
                    };
                })
                .Where(x => x.Count > 0)
                .ToList();

            if (rows.Count == 0)
            {
                return new List<RecommendationRow>();
            }

            double zValue = zValueOverride ?? InverseStandardNormalCdf((1.0 + 0.99) / 2.0);
            var avgRegression = BuildLinearRegression(rows.Select(r => (x: (double)r.Index, y: r.Avg)));
            var stdRegression = BuildLinearRegression(rows.Select(r => (x: (double)r.Index, y: Math.Max(r.StdDev, 0.0000001))));

            return rows.Select(r =>
            {
                double actualLower = r.Avg - (zValue * Math.Max(r.StdDev, 0.0000001));
                double actualUpper = r.Avg + (zValue * Math.Max(r.StdDev, 0.0000001));
                double trendAvg = avgRegression.Intercept + (avgRegression.Slope * r.Index);
                double trendStdDev = Math.Max(stdRegression.Intercept + (stdRegression.Slope * r.Index), 0.0000001);
                return new RecommendationRow
                {
                    ColumnName = r.Name,
                    Count = r.Count,
                    Avg = r.Avg,
                    StdDev = r.StdDev,
                    CurrentNgRatePercent = 0.0,
                    RequiredLowerActual = actualLower,
                    RequiredUpperActual = actualUpper,
                    TrendAvg = trendAvg,
                    TrendStdDev = trendStdDev,
                    RequiredLowerTrend = trendAvg - (zValue * trendStdDev),
                    RequiredUpperTrend = trendAvg + (zValue * trendStdDev)
                };
            }).ToList();
        }

        private string BuildRowLabel(int rowIndex)
        {
            return _result.XAxisType switch
            {
                XAxisValueType.Numeric => rowIndex < _result.XNumericValues.Count ? _result.XNumericValues[rowIndex].ToString("F4", CultureInfo.InvariantCulture) : (rowIndex + 1).ToString(CultureInfo.InvariantCulture),
                XAxisValueType.Date => rowIndex < _result.XDateValues.Count ? _result.XDateValues[rowIndex].ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : (rowIndex + 1).ToString(CultureInfo.InvariantCulture),
                _ => rowIndex < _result.XLabels.Count ? _result.XLabels[rowIndex] : (rowIndex + 1).ToString(CultureInfo.InvariantCulture)
            };
        }

        private double GetXAxisValue(int rowIndex)
        {
            return _result.XAxisType switch
            {
                XAxisValueType.Numeric => rowIndex < _result.XNumericValues.Count ? _result.XNumericValues[rowIndex] : rowIndex + 1,
                XAxisValueType.Date => OxyPlot.Axes.DateTimeAxis.ToDouble(rowIndex < _result.XDateValues.Count ? _result.XDateValues[rowIndex] : DateTime.MinValue.AddDays(rowIndex)),
                _ => rowIndex
            };
        }

        private static double ParseTargetNgRate(string text)
        {
            if (!GraphMakerParsingHelper.TryParseDouble(text, out double value))
            {
                value = 1.0;
            }

            return Math.Clamp(value, 0.0001, 99.0);
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
            return (slope, avgY - (slope * avgX));
        }

        private static double InverseStandardNormalCdf(double probability)
        {
            probability = Math.Clamp(probability, 0.0000001, 0.9999999);
            double[] a = { -39.69683028665376, 220.9460984245205, -275.9285104469687, 138.3577518672690, -30.66479806614716, 2.506628277459239 };
            double[] b = { -54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572 };
            double[] c = { -0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783 };
            double[] d = { 0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416 };
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

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
