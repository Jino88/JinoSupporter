using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using OxyPlot.Annotations;
using OxyPlot.Wpf;

namespace GraphMaker
{
    public partial class MultiColumnResultWindow : Window
    {
        private readonly MultiColumnGraphResult _result;
        private readonly string _xAxisName;
        private readonly string _fileName;

        public MultiColumnResultWindow(string fileName, string xAxisName, MultiColumnGraphResult result)
        {
            InitializeComponent();
            _fileName = fileName;
            _xAxisName = string.IsNullOrWhiteSpace(xAxisName) ? "X" : xAxisName;
            _result = result;

            HeaderTextBlock.Text = $"{fileName} - Multi Column Graphs ({result.Columns.Count} columns)";
            BuildOverallPanel();
            BuildPlotCards();
        }

        private void BuildOverallPanel()
        {
            OverallTotalRowsTextBox.Text = _result.Overall.TotalRows.ToString("N0", CultureInfo.InvariantCulture);
            OverallNgRowsTextBox.Text = _result.Overall.NgRows.ToString("N0", CultureInfo.InvariantCulture);
            OverallNgRateTextBox.Text = _result.Overall.NgRatePercent.ToString("F2", CultureInfo.InvariantCulture);
            double ppm = _result.Overall.NgRatePercent * 10000.0;
            OverallPpmTextBox.Text = ppm.ToString("F2", CultureInfo.InvariantCulture);
            RecommendedLimitsTextBox.Text = BuildRecommendedLimitsSummary();
        }

        private void BuildPlotCards()
        {
            PlotGrid.Children.Clear();
            PlotGrid.Columns = _result.Columns.Count <= 1 ? 1 : 2;

            foreach (ColumnGraphResult col in _result.Columns)
            {
                var cardBorder = new Border
                {
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(203, 213, 225)),
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Margin = new Thickness(4),
                    Padding = new Thickness(8)
                };

                var cardGrid = new Grid();
                cardGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                cardGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                cardGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var titleButton = new Button
                {
                    Content = col.ColumnName,
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Left,
                    Padding = new Thickness(0),
                    Margin = new Thickness(0, 0, 0, 6),
                    MinHeight = 0,
                    BorderThickness = new Thickness(0),
                    Background = System.Windows.Media.Brushes.Transparent,
                    Cursor = System.Windows.Input.Cursors.Hand,
                    ToolTip = "Click to open large view"
                };
                titleButton.Click += (_, _) => OpenSingleGraphWindow(col);
                Grid.SetRow(titleButton, 0);
                cardGrid.Children.Add(titleButton);

                var plotView = new PlotView
                {
                    MinHeight = _result.Columns.Count <= 1 ? 620 : 360,
                    Height = double.NaN,
                    VerticalAlignment = System.Windows.VerticalAlignment.Stretch,
                    Model = BuildPlotModel(col)
                };
                Grid.SetRow(plotView, 1);
                cardGrid.Children.Add(plotView);

                var infoBox = new TextBox
                {
                    Margin = new Thickness(0, 8, 0, 0),
                    IsReadOnly = true,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                    MinHeight = 96,
                    Text = BuildStatsText(col)
                };
                Grid.SetRow(infoBox, 2);
                cardGrid.Children.Add(infoBox);

                cardBorder.Child = cardGrid;
                PlotGrid.Children.Add(cardBorder);
            }
        }

        private PlotModel BuildPlotModel(ColumnGraphResult col)
        {
            var model = new PlotModel
            {
                Title = col.ColumnName,
                IsLegendVisible = true,
                Background = OxyColors.White,
                PlotAreaBackground = OxyColors.White
            };
            model.TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.TitleColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.PlotAreaBorderColor = OxyColor.FromRgb(0xB8, 0xC7, 0xDB);
            if (model.Legends.Count == 0)
            {
                model.Legends.Add(new Legend
                {
                    LegendPlacement = LegendPlacement.Outside,
                    LegendPosition = LegendPosition.RightTop,
                    LegendOrientation = LegendOrientation.Vertical
                });
            }

            switch (_result.XAxisType)
            {
                case XAxisValueType.Numeric:
                    model.Axes.Add(new LinearAxis
                    {
                        Position = AxisPosition.Bottom,
                        Title = _xAxisName,
                        AxislineStyle = LineStyle.Solid,
                        AxislineThickness = 1,
                        TickStyle = TickStyle.Outside,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.None
                    });
                    break;

                case XAxisValueType.Date:
                    model.Axes.Add(new DateTimeAxis
                    {
                        Position = AxisPosition.Bottom,
                        Title = _xAxisName,
                        StringFormat = "MM-dd",
                        AxislineStyle = LineStyle.Solid,
                        AxislineThickness = 1,
                        TickStyle = TickStyle.Outside,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.None
                    });
                    break;

                default:
                    int labelCount = _result.XLabels.Count;
                    int desiredTickCount = 10;
                    int labelStep = Math.Max(1, (int)Math.Ceiling(labelCount / (double)desiredTickCount));

                    var categoryAxis = new LinearAxis
                    {
                        Position = AxisPosition.Bottom,
                        Title = _xAxisName,
                        Minimum = 0,
                        Maximum = Math.Max(0, labelCount - 1),
                        MajorStep = labelStep,
                        MinorStep = double.NaN,
                        Angle = 35,
                        AxislineStyle = LineStyle.Solid,
                        AxislineThickness = 1,
                        TickStyle = TickStyle.Outside,
                        MajorGridlineStyle = LineStyle.Solid,
                        MinorGridlineStyle = LineStyle.None,
                        LabelFormatter = value =>
                        {
                            int idx = (int)Math.Round(value);
                            if (idx < 0 || idx >= labelCount)
                            {
                                return string.Empty;
                            }

                            if (Math.Abs(value - idx) > 0.01)
                            {
                                return string.Empty;
                            }

                            if (idx % labelStep != 0)
                            {
                                return string.Empty;
                            }

                            return _result.XLabels[idx];
                        }
                    };
                    model.Axes.Add(categoryAxis);
                    break;
            }

            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = col.ColumnName,
                AxislineStyle = LineStyle.Solid,
                AxislineThickness = 1,
                TickStyle = TickStyle.Outside,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.None
            });

            var series = new LineSeries
            {
                Title = "Data",
                RenderInLegend = true,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                StrokeThickness = 0,
                Color = OxyColors.SteelBlue,
                MarkerFill = OxyColors.SteelBlue
            };

            foreach (var row in col.Rows)
            {
                double x = row.RowIndex;
                switch (_result.XAxisType)
                {
                    case XAxisValueType.Numeric:
                        x = row.RowIndex < _result.XNumericValues.Count ? _result.XNumericValues[row.RowIndex] : row.RowIndex + 1;
                        break;
                    case XAxisValueType.Date:
                        DateTime dt = row.RowIndex < _result.XDateValues.Count ? _result.XDateValues[row.RowIndex] : DateTime.MinValue.AddDays(row.RowIndex);
                        x = DateTimeAxis.ToDouble(dt);
                        break;
                    case XAxisValueType.Category:
                        x = row.RowIndex;
                        break;
                }

                series.Points.Add(new DataPoint(x, row.Y));
            }

            model.Series.Add(series);
            AddAverageSeries(model, col);

            AddLimitLine(model, col.Upper, "Upper", OxyColors.IndianRed);
            AddLimitLine(model, col.Spec, "Spec", OxyColors.DarkOrange);
            AddLimitLine(model, col.Lower, "Lower", OxyColors.SeaGreen);

            return model;
        }

        private void AddAverageSeries(PlotModel model, ColumnGraphResult col)
        {
            var avgSeries = new LineSeries
            {
                Title = "Average",
                RenderInLegend = true,
                MarkerType = MarkerType.Diamond,
                MarkerSize = 3,
                StrokeThickness = 2,
                Color = OxyColors.DarkOrange,
                MarkerFill = OxyColors.DarkOrange
            };

            IEnumerable<IGrouping<double, ColumnGraphRow>> groups = col.Rows
                .GroupBy(row => GetXAxisValue(row.RowIndex))
                .OrderBy(group => group.Key);

            foreach (IGrouping<double, ColumnGraphRow> group in groups)
            {
                avgSeries.Points.Add(new DataPoint(group.Key, group.Average(item => item.Y)));
            }

            if (avgSeries.Points.Count > 0)
            {
                model.Series.Add(avgSeries);
            }
        }

        private double GetXAxisValue(int rowIndex)
        {
            return _result.XAxisType switch
            {
                XAxisValueType.Numeric => rowIndex < _result.XNumericValues.Count ? _result.XNumericValues[rowIndex] : rowIndex + 1,
                XAxisValueType.Date => DateTimeAxis.ToDouble(rowIndex < _result.XDateValues.Count ? _result.XDateValues[rowIndex] : DateTime.MinValue.AddDays(rowIndex)),
                _ => rowIndex
            };
        }

        private void OpenSingleGraphWindow(ColumnGraphResult col)
        {
            var window = new MultiColumnLargeDetailWindow(_fileName, col, _result, BuildPlotModel(col), BuildStatsText(col))
            {
                Owner = this
            };
            window.Show();
        }

        private static void AddLimitLine(PlotModel model, double? value, string title, OxyColor color)
        {
            if (!value.HasValue)
            {
                return;
            }

            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Horizontal,
                Y = value.Value,
                Color = color,
                Text = title,
                TextColor = color,
                StrokeThickness = 1.5,
                LineStyle = LineStyle.Dash
            });
        }

        private static string BuildStatsText(ColumnGraphResult col)
        {
            string cpkText = col.Cpk.HasValue ? col.Cpk.Value.ToString("F4", CultureInfo.InvariantCulture) : "-";
            string specText = col.Spec.HasValue ? col.Spec.Value.ToString("F6", CultureInfo.InvariantCulture) : "-";
            string upperText = col.Upper.HasValue ? col.Upper.Value.ToString("F6", CultureInfo.InvariantCulture) : "-";
            string lowerText = col.Lower.HasValue ? col.Lower.Value.ToString("F6", CultureInfo.InvariantCulture) : "-";
            double ppm = col.NgRatePercent * 10000.0;

            return string.Join(Environment.NewLine, new[]
            {
                $"Column: {col.ColumnName}",
                $"Spec / Upper / Lower: {specText} / {upperText} / {lowerText}",
                $"Total: {col.TotalCount:N0}",
                $"NG: {col.NgCount:N0}",
                $"NG Rate (%): {col.NgRatePercent:F2}",
                $"NG PPM: {ppm:F2}",
                $"Avg: {col.Avg:F6}",
                $"StdDev: {col.StdDev:F6}",
                $"CPK: {cpkText}"
            });
        }

        private string BuildRecommendedLimitsSummary()
        {
            if (_result.Columns.Count == 0)
            {
                return "No selected columns.";
            }

            const double targetNgRatePercent = 1.0;
            double yield = Math.Clamp(1.0 - (targetNgRatePercent / 100.0), 0.000001, 0.999999);
            double zValue = InverseStandardNormalCdf((1.0 + yield) / 2.0);

            double actualLower = _result.Columns.Min(c => c.Avg - (zValue * Math.Max(c.StdDev, 0.0000001)));
            double actualUpper = _result.Columns.Max(c => c.Avg + (zValue * Math.Max(c.StdDev, 0.0000001)));

            var avgRegression = BuildLinearRegression(_result.Columns.Select((column, index) => (x: (double)index, y: column.Avg)));
            var stdRegression = BuildLinearRegression(_result.Columns.Select((column, index) => (x: (double)index, y: Math.Max(column.StdDev, 0.0000001))));

            double trendLower = double.MaxValue;
            double trendUpper = double.MinValue;
            for (int index = 0; index < _result.Columns.Count; index++)
            {
                double trendAvg = avgRegression.Intercept + (avgRegression.Slope * index);
                double trendStdDev = Math.Max(stdRegression.Intercept + (stdRegression.Slope * index), 0.0000001);
                trendLower = Math.Min(trendLower, trendAvg - (zValue * trendStdDev));
                trendUpper = Math.Max(trendUpper, trendAvg + (zValue * trendStdDev));
            }

            string widestColumn = _result.Columns
                .OrderByDescending(c => 2.0 * zValue * Math.Max(c.StdDev, 0.0000001))
                .FirstOrDefault()?.ColumnName ?? "-";

            var sb = new StringBuilder();
            sb.AppendLine("Target NG Rate (%): 1.0000");
            sb.AppendLine($"Applied z-value: {zValue:F4}");
            sb.AppendLine(string.Empty);
            sb.AppendLine("[Actual]");
            sb.AppendLine($"LSL: {actualLower:F4}");
            sb.AppendLine($"USL: {actualUpper:F4}");
            sb.AppendLine(string.Empty);
            sb.AppendLine("[Trend]");
            sb.AppendLine($"LSL: {trendLower:F4}");
            sb.AppendLine($"USL: {trendUpper:F4}");
            sb.AppendLine(string.Empty);
            sb.AppendLine($"Widest Column: {widestColumn}");
            sb.Append("Rule: each selected column uses Avg +/- z * StdDev, then the outermost limits are shown.");
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
    }
}
