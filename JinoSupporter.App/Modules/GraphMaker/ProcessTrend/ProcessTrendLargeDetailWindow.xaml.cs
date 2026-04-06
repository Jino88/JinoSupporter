using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Annotations;
using OxyPlot.Legends;
using OxyPlot.Series;
using OxyPlot.Wpf;
using Color = System.Windows.Media.Color;
using Colors = System.Windows.Media.Colors;

namespace GraphMaker
{
    public partial class ProcessTrendLargeDetailWindow : Window
    {
        private sealed class SpecTargetOption
        {
            public string Key { get; init; } = string.Empty;
            public string Label { get; init; } = string.Empty;
        }

        private sealed class ElementStatRow
        {
            public string Element { get; init; } = string.Empty;
            public string MaxText { get; init; } = "-";
            public string MinText { get; init; } = "-";
            public string AvgText { get; init; } = "-";
            public string StdText { get; init; } = "-";
            public string CpkText { get; init; } = "-";
        }

        private sealed class RoutingReasonRow
        {
            public string Target { get; init; } = string.Empty;
            public string Routing { get; init; } = "-";
            public string Reason { get; init; } = "-";
        }

        private readonly ProcessPairPlotResult _pair;
        private readonly List<ProcessTrendComputationCandidate> _candidates = new();
        private readonly bool _useQuadraticRegression;
        private readonly Dictionary<string, bool> _seriesVisibility = new(StringComparer.Ordinal);

        public ProcessTrendLargeDetailWindow(ProcessPairPlotResult pair)
        {
            InitializeComponent();
            _pair = pair;
            PairTitleTextBlock.Text = pair.PairTitle;

            if (pair.ComputationCandidates.Count > 0)
            {
                _candidates.AddRange(pair.ComputationCandidates);
            }
            else
            {
                _candidates.Add(new ProcessTrendComputationCandidate
                {
                    PairTitle = pair.PairTitle,
                    PlotModel = pair.PlotModel,
                    XAxisTitle = pair.XAxisTitle,
                    YAxisTitle = pair.YAxisTitle,
                    RawPoints = pair.RawPoints
                });
            }

            _useQuadraticRegression = DetectQuadraticRegressionMode();
            PopulateSpecTargets();

            var allPoints = _candidates.SelectMany(c => c.RawPoints).ToList();
            double desiredDefault = allPoints.Count > 0 ? allPoints.Average(p => p.Y) : 430.0;
            DesiredYTextBox.Text = desiredDefault.ToString("F3", CultureInfo.InvariantCulture);
            DefectRateTextBox.Text = "1.0";
            BaseSummaryTextBlock.Text = BuildBaseSummary();
            UpdateModelAndSummary();
        }

        private void Input_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateModelAndSummary();
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void LegendColorsButton_Click(object sender, RoutedEventArgs e)
        {
            LargePlotWindowHelper.OpenLegendColorEditor(this, LargePlotView);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Select folder to save graph images"
            };

            if (dialog.ShowDialog(this) != true || string.IsNullOrWhiteSpace(dialog.FolderName))
            {
                return;
            }

            try
            {
                SavePlotModelAsPng(LargePlotView.Model, Path.Combine(dialog.FolderName, BuildSafeFileName($"{_pair.PairTitle}_Trend Graph.png")));
                SavePlotModelAsPng(NormalDistributionPlotView.Model, Path.Combine(dialog.FolderName, BuildSafeFileName($"{_pair.PairTitle}_Normal Distribution.png")));

                MessageBox.Show(this,
                    "Saved graph images as 1920x1080 PNG files.",
                    "Save Complete",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    $"Failed to save graph images: {ex.Message}",
                    "Save Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void SpecTargetComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            UpdateModelAndSummary();
        }

        private void DisplayModeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            UpdateModelAndSummary();
        }

        private void PopulateSpecTargets()
        {
            var items = new List<SpecTargetOption>
            {
                new() { Key = "combined", Label = $"{_pair.PairTitle} (Group total)" }
            };

            items.AddRange(_candidates.Select((candidate, index) => new SpecTargetOption
            {
                Key = $"candidate:{index}",
                Label = candidate.PairTitle
            }));

            SpecTargetComboBox.ItemsSource = items;
            SpecTargetComboBox.SelectedValue = "combined";
        }

        private void UpdateModelAndSummary()
        {
            if (!GraphMakerParsingHelper.TryParseDouble(DesiredYTextBox.Text, out double desiredY))
            {
                desiredY = 430.0;
            }

            if (!GraphMakerParsingHelper.TryParseDouble(DefectRateTextBox.Text, out double defectRatePct))
            {
                defectRatePct = 1.0;
            }

            defectRatePct = Math.Clamp(defectRatePct, 0.0001, 49.9999);
            var model = BuildModel(desiredY, defectRatePct);
            CaptureDefaultSeriesVisibility(model);
            ApplySeriesVisibility(model);
            LargePlotView.Model = model;
            NormalDistributionPlotView.Model = BuildNormalDistributionModel(model);
            ElementStatsGrid.ItemsSource = BuildElementStatRows(model);
            RoutingReasonGrid.ItemsSource = BuildRoutingReasonRows(desiredY, defectRatePct);
            PopulateSeriesVisibilityPanel();
            ComputedSummaryTextBlock.Text = BuildComputedSummary(desiredY, defectRatePct);
        }

        private PlotModel BuildModel(double desiredY, double defectRatePct)
        {
            var model = IsCombinedMode()
                ? BuildCombinedModel()
                : ClonePlotModel(_pair.PlotModel);

            if (TryGetCategoryAxis(_pair.PlotModel, AxisPosition.Bottom, out CategoryAxis? sourceCategoryAxis))
            {
                ForceCategoryXAxis(model, sourceCategoryAxis);
            }

            EnsureDetailLegend(model);
            RemoveDynamicGuideLines(model);
            return model;
        }

        private PlotModel BuildCombinedModel()
        {
            bool useCategoryXAxis = _pair.PlotModel.Axes.Any(axis => axis.Position == AxisPosition.Bottom && axis is CategoryAxis);
            var model = CreateBaseModelFromSourceAxes(_pair.PairTitle, _pair.PlotModel);
            var allPoints = _candidates.SelectMany(c => c.RawPoints).ToList();
            var combinedColor = GetCombinedSeriesColor();

            Series combinedSeries = useCategoryXAxis
                ? new LineSeries
                {
                    Title = $"{_pair.PairTitle} Combined",
                    Color = combinedColor,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3,
                    MarkerFill = combinedColor,
                    MarkerStroke = combinedColor,
                    StrokeThickness = 0,
                    LineStyle = LineStyle.None,
                    RenderInLegend = true
                }
                : new ScatterSeries
                {
                    Title = $"{_pair.PairTitle} Combined",
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3,
                    MarkerFill = combinedColor,
                    MarkerStroke = combinedColor,
                    RenderInLegend = true
                };

            foreach (var point in allPoints)
            {
                if (combinedSeries is LineSeries lineSeries)
                {
                    lineSeries.Points.Add(new DataPoint(point.X, point.Y));
                }
                else if (combinedSeries is ScatterSeries scatterSeries)
                {
                    scatterSeries.Points.Add(new ScatterPoint(point.X, point.Y));
                }
            }

            model.Series.Add(combinedSeries);
            AddTrendSeries(model, $"{_pair.PairTitle} Regression", allPoints, combinedColor, true);
            CopySpecLines(model, _pair.PlotModel);
            CopyAnnotations(model, _pair.PlotModel);
            return model;
        }

        private static void CopySpecLines(PlotModel target, PlotModel source)
        {
            foreach (var line in source.Series.OfType<LineSeries>())
            {
                if (string.IsNullOrWhiteSpace(line.Title) ||
                    (!line.Title.EndsWith(" USL", StringComparison.OrdinalIgnoreCase) &&
                     !line.Title.EndsWith(" LSL", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                var cloned = new LineSeries
                {
                    Title = line.Title,
                    Color = line.Color,
                    StrokeThickness = line.StrokeThickness,
                    LineStyle = line.LineStyle,
                    MarkerType = line.MarkerType,
                    MarkerSize = line.MarkerSize,
                    MarkerFill = line.MarkerFill,
                    RenderInLegend = line.RenderInLegend
                };

                foreach (var point in line.Points)
                {
                    cloned.Points.Add(point);
                }

                target.Series.Add(cloned);
            }
        }

        private static void EnsureDetailLegend(PlotModel model)
        {
            if (model.Legends.Count > 0)
            {
                return;
            }

            model.Legends.Add(new Legend
            {
                LegendPlacement = LegendPlacement.Inside,
                LegendPosition = LegendPosition.LeftTop,
                LegendOrientation = LegendOrientation.Vertical,
                LegendBackground = OxyColor.FromAColor(220, OxyColors.White)
            });
        }

        private static void RemoveDynamicGuideLines(PlotModel model)
        {
            var removableTitles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Desired Y",
                "X@Desired Y"
            };

            for (int i = model.Series.Count - 1; i >= 0; i--)
            {
                if (model.Series[i] is not LineSeries ls || string.IsNullOrWhiteSpace(ls.Title))
                {
                    continue;
                }

                if (ls.Title.StartsWith("Desired Y=", StringComparison.OrdinalIgnoreCase) ||
                    ls.Title.StartsWith("X@", StringComparison.OrdinalIgnoreCase) ||
                    removableTitles.Contains(ls.Title))
                {
                    model.Series.RemoveAt(i);
                }
            }
        }

        private PlotModel CreateBaseModel(string title, string xAxisTitle, string yAxisTitle)
        {
            var model = new PlotModel
            {
                Title = title,
                Background = OxyColors.White,
                PlotAreaBackground = OxyColors.White,
                TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33),
                TitleColor = OxyColor.FromRgb(0x1D, 0x27, 0x33),
                PlotAreaBorderColor = OxyColor.FromRgb(0xB8, 0xC7, 0xDB),
                IsLegendVisible = true
            };

            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = xAxisTitle,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = yAxisTitle,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            EnsureDetailLegend(model);
            return model;
        }

        private static bool TryGetCategoryAxis(PlotModel model, AxisPosition position, out CategoryAxis? categoryAxis)
        {
            categoryAxis = model.Axes.OfType<CategoryAxis>().FirstOrDefault(axis => axis.Position == position);
            return categoryAxis != null;
        }

        private static void ForceCategoryXAxis(PlotModel model, CategoryAxis sourceCategoryAxis)
        {
            for (int i = model.Axes.Count - 1; i >= 0; i--)
            {
                if (model.Axes[i].Position == AxisPosition.Bottom)
                {
                    model.Axes.RemoveAt(i);
                }
            }

            var categoryAxis = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                Title = sourceCategoryAxis.Title,
                GapWidth = sourceCategoryAxis.GapWidth,
                Angle = sourceCategoryAxis.Angle,
                MajorGridlineStyle = sourceCategoryAxis.MajorGridlineStyle,
                MinorGridlineStyle = sourceCategoryAxis.MinorGridlineStyle
            };

            foreach (string label in sourceCategoryAxis.Labels)
            {
                categoryAxis.Labels.Add(label);
            }

            model.Axes.Insert(0, categoryAxis);

            for (int i = 0; i < model.Series.Count; i++)
            {
                if (model.Series[i] is not ScatterSeries scatter)
                {
                    continue;
                }

                var lineSeries = new LineSeries
                {
                    Title = scatter.Title,
                    Color = scatter.MarkerStroke.IsUndefined() ? scatter.MarkerFill : scatter.MarkerStroke,
                    MarkerType = scatter.MarkerType,
                    MarkerSize = scatter.MarkerSize,
                    MarkerFill = scatter.MarkerFill,
                    MarkerStroke = scatter.MarkerStroke,
                    MarkerStrokeThickness = scatter.MarkerStrokeThickness,
                    StrokeThickness = 0,
                    LineStyle = LineStyle.None,
                    RenderInLegend = scatter.RenderInLegend
                };

                foreach (var point in scatter.Points)
                {
                    lineSeries.Points.Add(new DataPoint(point.X, point.Y));
                }

                model.Series[i] = lineSeries;
            }
        }

        private PlotModel CreateBaseModelFromSourceAxes(string title, PlotModel source)
        {
            var model = new PlotModel
            {
                Title = title,
                Background = OxyColors.White,
                PlotAreaBackground = OxyColors.White,
                TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33),
                TitleColor = OxyColor.FromRgb(0x1D, 0x27, 0x33),
                PlotAreaBorderColor = OxyColor.FromRgb(0xB8, 0xC7, 0xDB),
                IsLegendVisible = true
            };

            foreach (var axis in source.Axes)
            {
                if (axis is LinearAxis la)
                {
                    model.Axes.Add(new LinearAxis
                    {
                        Position = la.Position,
                        Title = la.Title,
                        MajorGridlineStyle = la.MajorGridlineStyle,
                        MinorGridlineStyle = la.MinorGridlineStyle,
                        Minimum = double.IsNaN(la.ActualMinimum) ? double.NaN : la.ActualMinimum,
                        Maximum = double.IsNaN(la.ActualMaximum) ? double.NaN : la.ActualMaximum
                    });
                }
                else if (axis is CategoryAxis ca)
                {
                    var categoryAxis = new CategoryAxis
                    {
                        Position = ca.Position,
                        Title = ca.Title,
                        GapWidth = ca.GapWidth,
                        Angle = ca.Angle,
                        MajorGridlineStyle = ca.MajorGridlineStyle,
                        MinorGridlineStyle = ca.MinorGridlineStyle
                    };

                    foreach (string label in ca.Labels)
                    {
                        categoryAxis.Labels.Add(label);
                    }

                    model.Axes.Add(categoryAxis);
                }
            }

            EnsureDetailLegend(model);
            return model;
        }

        private void AddTrendSeries(PlotModel model, string title, IReadOnlyList<DataPoint> points, OxyColor color, bool renderInLegend)
        {
            if (points.Count < 2)
            {
                return;
            }

            double minX = points.Min(p => p.X);
            double maxX = points.Max(p => p.X);
            if (Math.Abs(maxX - minX) <= 1e-12)
            {
                return;
            }

            if (!_useQuadraticRegression && TryCalculateTrendLine(points, out double slope, out double intercept, out _, out _))
            {
                var trend = new LineSeries
                {
                    Title = title,
                    StrokeThickness = 2,
                    Color = color,
                    RenderInLegend = renderInLegend
                };
                trend.Points.Add(new DataPoint(minX, slope * minX + intercept));
                trend.Points.Add(new DataPoint(maxX, slope * maxX + intercept));
                model.Series.Add(trend);
                return;
            }

            if (_useQuadraticRegression &&
                TryCalculateQuadraticTrendLine(points, out double a, out double b, out double c))
            {
                var trend = new LineSeries
                {
                    Title = title,
                    StrokeThickness = 2,
                    Color = color,
                    RenderInLegend = renderInLegend
                };

                const int segments = 60;
                for (int i = 0; i <= segments; i++)
                {
                    double x = minX + (maxX - minX) * i / segments;
                    double y = a * x * x + b * x + c;
                    trend.Points.Add(new DataPoint(x, y));
                }

                model.Series.Add(trend);
            }
        }

        private static PlotModel ClonePlotModel(PlotModel source)
        {
            var clone = new PlotModel
            {
                Title = source.Title,
                Subtitle = source.Subtitle,
                Background = source.Background,
                PlotAreaBackground = source.PlotAreaBackground,
                PlotAreaBorderColor = source.PlotAreaBorderColor,
                TextColor = source.TextColor,
                TitleColor = source.TitleColor,
                IsLegendVisible = source.IsLegendVisible
            };

            foreach (var axis in source.Axes)
            {
                if (axis is LinearAxis la)
                {
                    clone.Axes.Add(new LinearAxis
                    {
                        Position = la.Position,
                        Title = la.Title,
                        MajorGridlineStyle = la.MajorGridlineStyle,
                        MinorGridlineStyle = la.MinorGridlineStyle
                    });
                }
                else if (axis is CategoryAxis ca)
                {
                    var categoryAxis = new CategoryAxis
                    {
                        Position = ca.Position,
                        Title = ca.Title,
                        GapWidth = ca.GapWidth,
                        Angle = ca.Angle,
                        MajorGridlineStyle = ca.MajorGridlineStyle,
                        MinorGridlineStyle = ca.MinorGridlineStyle
                    };

                    foreach (string label in ca.Labels)
                    {
                        categoryAxis.Labels.Add(label);
                    }

                    clone.Axes.Add(categoryAxis);
                }
            }

            foreach (var series in source.Series)
            {
                switch (series)
                {
                    case ScatterSeries ss:
                        var newSs = new ScatterSeries
                        {
                            Title = ss.Title,
                            MarkerType = ss.MarkerType,
                            MarkerSize = ss.MarkerSize,
                            MarkerFill = ss.MarkerFill,
                            MarkerStroke = ss.MarkerStroke,
                            MarkerStrokeThickness = ss.MarkerStrokeThickness,
                            RenderInLegend = ss.RenderInLegend
                        };
                        foreach (var p in ss.Points)
                        {
                            newSs.Points.Add(new ScatterPoint(p.X, p.Y, p.Size, p.Value, p.Tag));
                        }
                        clone.Series.Add(newSs);
                        break;

                    case LineSeries ls:
                        var newLs = new LineSeries
                        {
                            Title = ls.Title,
                            Color = ls.Color,
                            StrokeThickness = ls.StrokeThickness,
                            LineStyle = ls.LineStyle,
                            MarkerType = ls.MarkerType,
                            MarkerSize = ls.MarkerSize,
                            MarkerFill = ls.MarkerFill,
                            RenderInLegend = ls.RenderInLegend
                        };
                        foreach (var p in ls.Points)
                        {
                            newLs.Points.Add(p);
                        }
                        clone.Series.Add(newLs);
                        break;
                }
            }

            CopyAnnotations(clone, source);

            return clone;
        }

        private static void CopyAnnotations(PlotModel target, PlotModel source)
        {
            foreach (var annotation in source.Annotations)
            {
                switch (annotation)
                {
                    case PolygonAnnotation polygon:
                        var polygonClone = new PolygonAnnotation
                        {
                            Layer = polygon.Layer,
                            Fill = polygon.Fill,
                            Stroke = polygon.Stroke,
                            StrokeThickness = polygon.StrokeThickness,
                            LineStyle = polygon.LineStyle,
                            ToolTip = polygon.ToolTip,
                            Text = polygon.Text,
                            TextColor = polygon.TextColor
                        };

                        foreach (var point in polygon.Points)
                        {
                            polygonClone.Points.Add(point);
                        }

                        target.Annotations.Add(polygonClone);
                        break;
                }
            }
        }

        private PlotModel BuildNormalDistributionModel(PlotModel source)
        {
            var model = CreateBaseModel($"{_pair.PairTitle} Normal Distribution", "Value", "Density");
            var seriesIndex = 1;

            foreach (var series in source.Series)
            {
                if (!ShouldIncludeInNormalDistribution(series))
                {
                    continue;
                }

                var values = ExtractSeriesValues(series);
                if (values.Count < 2)
                {
                    continue;
                }

                double mean = values.Average();
                double variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
                double stdDev = Math.Sqrt(variance);
                if (stdDev <= 1e-12)
                {
                    continue;
                }

                double min = values.Min() - stdDev;
                double max = values.Max() + stdDev;
                var distributionSeries = new LineSeries
                {
                    Title = string.IsNullOrWhiteSpace(series.Title) ? $"Series {seriesIndex}" : series.Title,
                    StrokeThickness = 2,
                    RenderInLegend = true,
                    Color = GetSeriesColor(series)
                };

                const int segments = 100;
                for (int i = 0; i <= segments; i++)
                {
                    double x = min + ((max - min) * i / segments);
                    double y = (1.0 / (stdDev * Math.Sqrt(2 * Math.PI))) *
                               Math.Exp(-0.5 * Math.Pow((x - mean) / stdDev, 2));
                    distributionSeries.Points.Add(new DataPoint(x, y));
                }

                model.Series.Add(distributionSeries);
                seriesIndex++;
            }

            return model;
        }

        private void CaptureDefaultSeriesVisibility(PlotModel model)
        {
            foreach (var series in model.Series)
            {
                var key = GetSeriesKey(series);
                if (!_seriesVisibility.ContainsKey(key))
                {
                    _seriesVisibility[key] = true;
                }
            }
        }

        private void ApplySeriesVisibility(PlotModel model)
        {
            for (int i = model.Series.Count - 1; i >= 0; i--)
            {
                var key = GetSeriesKey(model.Series[i]);
                if (_seriesVisibility.TryGetValue(key, out bool isVisible) && !isVisible)
                {
                    model.Series.RemoveAt(i);
                }
            }
        }

        private void PopulateSeriesVisibilityPanel()
        {
            SeriesVisibilityPanel.Children.Clear();

            if (LargePlotView.Model == null)
            {
                return;
            }

            foreach (var pair in _seriesVisibility
                         .OrderBy(entry => GetSeriesVisibilityOrder(entry.Key))
                         .ThenBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase))
            {
                var checkBox = new CheckBox
                {
                    Content = pair.Key,
                    IsChecked = pair.Value,
                    Margin = new Thickness(0, 0, 0, 6)
                };

                string capturedKey = pair.Key;
                checkBox.Checked += (_, _) => SetSeriesVisibility(capturedKey, true);
                checkBox.Unchecked += (_, _) => SetSeriesVisibility(capturedKey, false);
                SeriesVisibilityPanel.Children.Add(checkBox);
            }
        }

        private void SetSeriesVisibility(string seriesKey, bool isVisible)
        {
            _seriesVisibility[seriesKey] = isVisible;
            UpdateModelAndSummary();
        }

        private static int GetSeriesVisibilityOrder(string seriesKey)
        {
            if (seriesKey.EndsWith(" LSL", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }

            if (seriesKey.EndsWith(" USL", StringComparison.OrdinalIgnoreCase))
            {
                return 4;
            }

            if (seriesKey.EndsWith(" SPEC", StringComparison.OrdinalIgnoreCase))
            {
                return 5;
            }

            return 0;
        }

        private static bool ShouldIncludeInNormalDistribution(Series series)
        {
            if (string.IsNullOrWhiteSpace(series.Title))
            {
                return false;
            }

            return !series.Title.EndsWith(" USL", StringComparison.OrdinalIgnoreCase) &&
                   !series.Title.EndsWith(" LSL", StringComparison.OrdinalIgnoreCase) &&
                   !series.Title.Contains("Regression", StringComparison.OrdinalIgnoreCase);
        }

        private static List<double> ExtractSeriesValues(Series series)
        {
            return series switch
            {
                LineSeries line => line.Points
                    .Select(point => point.Y)
                    .Where(value => !double.IsNaN(value) && !double.IsInfinity(value))
                    .ToList(),
                ScatterSeries scatter => scatter.Points
                    .Select(point => point.Y)
                    .Where(value => !double.IsNaN(value) && !double.IsInfinity(value))
                    .ToList(),
                _ => new List<double>()
            };
        }

        private static OxyColor GetSeriesColor(Series series)
        {
            return series switch
            {
                LineSeries line => line.Color.IsUndefined() ? line.MarkerFill : line.Color,
                ScatterSeries scatter => scatter.MarkerFill,
                _ => OxyColors.Black
            };
        }

        private static string GetSeriesKey(Series series)
        {
            return string.IsNullOrWhiteSpace(series.Title) ? "Unnamed Series" : series.Title;
        }

        private IReadOnlyList<ElementStatRow> BuildElementStatRows(PlotModel source)
        {
            var rows = new List<ElementStatRow>();

            foreach (var series in source.Series.Where(ShouldIncludeInNormalDistribution))
            {
                string title = series.Title ?? "Unnamed Series";
                var values = ExtractSeriesValues(series);
                if (values.Count == 0)
                {
                    continue;
                }

                double max = values.Max();
                double min = values.Min();
                double avg = values.Average();
                double std = CalculateSampleStdDev(values, avg);
                double? cpk = TryCalculateSeriesCpk(source, title, avg, std);

                rows.Add(new ElementStatRow
                {
                    Element = title,
                    MaxText = max.ToString("F4", CultureInfo.InvariantCulture),
                    MinText = min.ToString("F4", CultureInfo.InvariantCulture),
                    AvgText = avg.ToString("F4", CultureInfo.InvariantCulture),
                    StdText = std.ToString("F4", CultureInfo.InvariantCulture),
                    CpkText = cpk.HasValue ? cpk.Value.ToString("F4", CultureInfo.InvariantCulture) : "-"
                });
            }

            return rows;
        }

        private static double? TryCalculateSeriesCpk(PlotModel source, string seriesTitle, double avg, double std)
        {
            if (std <= 1e-12)
            {
                return null;
            }

            double? usl = null;
            double? lsl = null;
            string uslTitle = $"{seriesTitle} USL";
            string lslTitle = $"{seriesTitle} LSL";

            foreach (LineSeries line in source.Series.OfType<LineSeries>())
            {
                if (line.Points.Count == 0 || string.IsNullOrWhiteSpace(line.Title))
                {
                    continue;
                }

                if (string.Equals(line.Title, uslTitle, StringComparison.OrdinalIgnoreCase))
                {
                    usl = line.Points[0].Y;
                }
                else if (string.Equals(line.Title, lslTitle, StringComparison.OrdinalIgnoreCase))
                {
                    lsl = line.Points[0].Y;
                }
            }

            if (!usl.HasValue && !lsl.HasValue)
            {
                return null;
            }

            double? cpu = usl.HasValue ? (usl.Value - avg) / (3d * std) : null;
            double? cpl = lsl.HasValue ? (avg - lsl.Value) / (3d * std) : null;

            return (cpu, cpl) switch
            {
                ({ } upper, { } lower) => Math.Min(upper, lower),
                ({ } upper, null) => upper,
                (null, { } lower) => lower,
                _ => null
            };
        }


        private string BuildBaseSummary()
        {
            var sb = new StringBuilder();

            foreach (var target in GetSelectedSpecTargets())
            {
                AppendSummaryForCandidate(sb, target.Title, target.Points);
            }

            return sb.ToString().TrimEnd();
        }

        private string BuildComputedSummary(double desiredY, double defectRatePct)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Desired(Y mean) = {desiredY:F4}");
            sb.AppendLine($"Defect rate target = {defectRatePct:F4}%");
            sb.AppendLine();

            foreach (var target in GetSelectedSpecTargets())
            {
                AppendComputedSummaryForCandidate(sb, target.Title, target.Points, desiredY, defectRatePct);
            }

            return sb.ToString().TrimEnd();
        }

        private IReadOnlyList<RoutingReasonRow> BuildRoutingReasonRows(double desiredY, double defectRatePct)
        {
            var rows = new List<RoutingReasonRow>();

            foreach (var target in GetSelectedSpecTargets())
            {
                rows.Add(BuildRoutingReasonRow(target.Title, target.Points, desiredY, defectRatePct));
            }

            return rows;
        }

        private IReadOnlyList<(string Title, IReadOnlyList<DataPoint> Points)> GetSelectedSpecTargets()
        {
            if (SpecTargetComboBox.SelectedValue is string key)
            {
                if (string.Equals(key, "combined", StringComparison.Ordinal))
                {
                    return new[]
                    {
                        (Title: $"{_pair.PairTitle} Combined", Points: (IReadOnlyList<DataPoint>)_candidates.SelectMany(c => c.RawPoints).ToList())
                    };
                }

                if (key.StartsWith("candidate:", StringComparison.Ordinal) &&
                    int.TryParse(key.Substring("candidate:".Length), out int idx) &&
                    idx >= 0 && idx < _candidates.Count)
                {
                    var candidate = _candidates[idx];
                    return new[]
                    {
                        (Title: candidate.PairTitle, Points: candidate.RawPoints)
                    };
                }
            }

            return new[]
            {
                (Title: $"{_pair.PairTitle} Combined", Points: (IReadOnlyList<DataPoint>)_candidates.SelectMany(c => c.RawPoints).ToList())
            };
        }

        private void AppendSummaryForCandidate(StringBuilder sb, string title, IReadOnlyList<DataPoint> points)
        {
            sb.AppendLine($"[{title}] n={points.Count:N0}");
            AppendBasicStats(sb, points);
            if (TryCalculateTrendLine(points, out double slope, out double intercept, out double r2, out double sigma))
            {
                sb.AppendLine($"Formula: Y = {slope:F6} * X + {intercept:F6}");
                sb.AppendLine($"R^2 = {r2:F4}, Residual sigma = {sigma:F4}");
            }
            else
            {
                sb.AppendLine("Regression unavailable");
            }
            sb.AppendLine();
        }

        private void AppendComputedSummaryForCandidate(
            StringBuilder sb,
            string title,
            IReadOnlyList<DataPoint> points,
            double desiredY,
            double defectRatePct)
        {
            sb.AppendLine($"[{title}]");
            if (TryCalculateTrendLine(points, out double slope, out double intercept, out double rSquared, out double sigma) &&
                Math.Abs(slope) > 1e-12)
            {
                AppendBasicStats(sb, points);
                double xRange = points.Count > 0 ? points.Max(p => p.X) - points.Min(p => p.X) : 0;
                double yRangeFromSlope = Math.Abs(slope) * xRange;
                double z = InverseStandardNormalCdf(1.0 - defectRatePct / 100.0);
                double requiredMeanY = desiredY + z * sigma;

                sb.AppendLine($"Formula = Y = {slope:F6} * X + {intercept:F6}");
                sb.AppendLine($"R^2 = {rSquared:F4}");
                sb.AppendLine($"Residual sigma = {sigma:F4}");
                sb.AppendLine($"Z value = {z:F4}");
                sb.AppendLine($"Required mean Y (for defect) = {requiredMeanY:F4}");

                if (Math.Abs(slope) < 0.001 || rSquared < 0.05 || yRangeFromSlope < sigma * 0.5)
                {
                    sb.AppendLine("Required X calculation is unstable.");
                    sb.AppendLine($"Reason: slope is too small or fit is weak.");
                    sb.AppendLine($"X range = {xRange:F4}, |slope|*X range = {yRangeFromSlope:F4}");
                }
                else
                {
                    double xForDefect = (requiredMeanY - intercept) / slope;
                    double xForDesired = (desiredY - intercept) / slope;
                    sb.AppendLine($"Required min X (defect) = {xForDefect:F4}");
                    sb.AppendLine($"Required X (desired Y) = {xForDesired:F4}");
                }
            }
            else
            {
                sb.AppendLine("Required X cannot be computed.");
            }
            sb.AppendLine();
        }

        private static RoutingReasonRow BuildRoutingReasonRow(
            string title,
            IReadOnlyList<DataPoint> points,
            double desiredY,
            double defectRatePct)
        {
            if (!TryCalculateTrendLine(points, out double slope, out double intercept, out double rSquared, out double sigma))
            {
                return new RoutingReasonRow
                {
                    Target = title,
                    Routing = "Unavailable",
                    Reason = "Regression unavailable"
                };
            }

            if (Math.Abs(slope) <= 1e-12)
            {
                return new RoutingReasonRow
                {
                    Target = title,
                    Routing = "Unavailable",
                    Reason = "Slope is zero"
                };
            }

            double xRange = points.Count > 0 ? points.Max(p => p.X) - points.Min(p => p.X) : 0;
            double yRangeFromSlope = Math.Abs(slope) * xRange;
            if (Math.Abs(slope) < 0.001 || rSquared < 0.05 || yRangeFromSlope < sigma * 0.5)
            {
                return new RoutingReasonRow
                {
                    Target = title,
                    Routing = "Unstable",
                    Reason = "Slope is too small or fit is weak"
                };
            }

            double z = InverseStandardNormalCdf(1.0 - defectRatePct / 100.0);
            double requiredMeanY = desiredY + z * sigma;
            double xForDefect = (requiredMeanY - intercept) / slope;

            return new RoutingReasonRow
            {
                Target = title,
                Routing = xForDefect.ToString("F4", CultureInfo.InvariantCulture),
                Reason = "Calculated from regression"
            };
        }

        private static void AppendBasicStats(StringBuilder sb, IReadOnlyList<DataPoint> points)
        {
            if (points.Count == 0)
            {
                return;
            }

            double xMean = points.Average(p => p.X);
            double yMean = points.Average(p => p.Y);
            double xStd = CalculateSampleStdDev(points.Select(p => p.X).ToList(), xMean);
            double yStd = CalculateSampleStdDev(points.Select(p => p.Y).ToList(), yMean);

            sb.AppendLine($"X mean = {xMean:F4}, X std = {xStd:F4}");
            sb.AppendLine($"Y mean = {yMean:F4}, Y std = {yStd:F4}");
        }

        private static double CalculateSampleStdDev(IReadOnlyList<double> values, double mean)
        {
            if (values.Count < 2)
            {
                return 0;
            }

            double sum = 0;
            foreach (double value in values)
            {
                double diff = value - mean;
                sum += diff * diff;
            }

            return Math.Sqrt(sum / (values.Count - 1));
        }

        private static bool TryCalculateTrendLine(
            IReadOnlyList<DataPoint> points,
            out double slope,
            out double intercept,
            out double rSquared,
            out double sigma)
        {
            slope = 0;
            intercept = 0;
            rSquared = 0;
            sigma = 0;

            if (points.Count < 3)
            {
                return false;
            }

            double n = points.Count;
            double sumX = points.Sum(p => p.X);
            double sumY = points.Sum(p => p.Y);
            double sumXY = points.Sum(p => p.X * p.Y);
            double sumXX = points.Sum(p => p.X * p.X);
            double denominator = n * sumXX - sumX * sumX;
            if (Math.Abs(denominator) < 1e-12)
            {
                return false;
            }

            slope = (n * sumXY - sumX * sumY) / denominator;
            intercept = (sumY - slope * sumX) / n;
            double meanY = sumY / n;

            double ssTot = 0;
            double ssRes = 0;
            foreach (var p in points)
            {
                double predicted = slope * p.X + intercept;
                ssTot += Math.Pow(p.Y - meanY, 2);
                ssRes += Math.Pow(p.Y - predicted, 2);
            }
            rSquared = ssTot <= 1e-12 ? 1.0 : 1.0 - (ssRes / ssTot);
            sigma = Math.Sqrt(ssRes / (points.Count - 2));
            return true;
        }

        private static bool TryCalculateQuadraticTrendLine(
            IReadOnlyList<DataPoint> points,
            out double a,
            out double b,
            out double c)
        {
            a = 0;
            b = 0;
            c = 0;

            if (points.Count < 3)
            {
                return false;
            }

            double sumX = 0;
            double sumX2 = 0;
            double sumX3 = 0;
            double sumX4 = 0;
            double sumY = 0;
            double sumXY = 0;
            double sumX2Y = 0;

            foreach (var point in points)
            {
                double x = point.X;
                double y = point.Y;
                double x2 = x * x;
                sumX += x;
                sumX2 += x2;
                sumX3 += x2 * x;
                sumX4 += x2 * x2;
                sumY += y;
                sumXY += x * y;
                sumX2Y += x2 * y;
            }

            double n = points.Count;
            double[,] matrix =
            {
                { sumX4, sumX3, sumX2, sumX2Y },
                { sumX3, sumX2, sumX, sumXY },
                { sumX2, sumX, n, sumY }
            };

            return Solve3x3(matrix, out a, out b, out c);
        }

        private static bool Solve3x3(double[,] matrix, out double x, out double y, out double z)
        {
            x = 0;
            y = 0;
            z = 0;

            const int size = 3;
            for (int pivot = 0; pivot < size; pivot++)
            {
                int maxRow = pivot;
                for (int row = pivot + 1; row < size; row++)
                {
                    if (Math.Abs(matrix[row, pivot]) > Math.Abs(matrix[maxRow, pivot]))
                    {
                        maxRow = row;
                    }
                }

                if (Math.Abs(matrix[maxRow, pivot]) < 1e-12)
                {
                    return false;
                }

                if (maxRow != pivot)
                {
                    for (int col = pivot; col <= size; col++)
                    {
                        (matrix[pivot, col], matrix[maxRow, col]) = (matrix[maxRow, col], matrix[pivot, col]);
                    }
                }

                double divisor = matrix[pivot, pivot];
                for (int col = pivot; col <= size; col++)
                {
                    matrix[pivot, col] /= divisor;
                }

                for (int row = 0; row < size; row++)
                {
                    if (row == pivot)
                    {
                        continue;
                    }

                    double factor = matrix[row, pivot];
                    for (int col = pivot; col <= size; col++)
                    {
                        matrix[row, col] -= factor * matrix[pivot, col];
                    }
                }
            }

            x = matrix[0, 3];
            y = matrix[1, 3];
            z = matrix[2, 3];
            return true;
        }

        private bool DetectQuadraticRegressionMode()
        {
            return _candidates
                .SelectMany(c => c.PlotModel.Series.OfType<LineSeries>())
                .Any(series => !string.IsNullOrWhiteSpace(series.Title) &&
                               series.Title.Contains("Regression", StringComparison.OrdinalIgnoreCase) &&
                               series.Points.Count > 2);
        }

        private bool IsCombinedMode()
        {
            return DisplayModeComboBox.SelectedIndex == 1;
        }

        private OxyColor GetCombinedSeriesColor()
        {
            if (_candidates.Count == 0)
            {
                return OxyColors.Black;
            }

            var seedSeries = _candidates[0].PlotModel.Series
                .OfType<ScatterSeries>()
                .FirstOrDefault();
            if (seedSeries != null)
            {
                return seedSeries.MarkerFill;
            }

            return OxyColor.FromArgb(Colors.DodgerBlue.A, Colors.DodgerBlue.R, Colors.DodgerBlue.G, Colors.DodgerBlue.B);
        }

        private static void SavePlotModelAsPng(PlotModel? model, string path)
        {
            if (model == null)
            {
                return;
            }

            var exporter = new PngExporter
            {
                Width = 1920,
                Height = 1080
            };
            exporter.ExportToFile(model, path);
        }

        private static string BuildSafeFileName(string fileName)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new string(fileName.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray());
            return sanitized;
        }

        // Acklam's inverse normal CDF approximation.
        private static double InverseStandardNormalCdf(double p)
        {
            p = Math.Clamp(p, 1e-12, 1 - 1e-12);

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

            const double plow = 0.02425;
            const double phigh = 1 - plow;

            if (p < plow)
            {
                double q = Math.Sqrt(-2 * Math.Log(p));
                return (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                       ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
            }

            if (p > phigh)
            {
                double q = Math.Sqrt(-2 * Math.Log(1 - p));
                return -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) /
                        ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
            }

            double r = p - 0.5;
            double s = r * r;
            return (((((a[0] * s + a[1]) * s + a[2]) * s + a[3]) * s + a[4]) * s + a[5]) * r /
                   (((((b[0] * s + b[1]) * s + b[2]) * s + b[3]) * s + b[4]) * s + 1);
        }
    }
}
