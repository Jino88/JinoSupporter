using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;
using OxyPlot.Wpf;
using Color = System.Windows.Media.Color;
using MessageBox = System.Windows.MessageBox;

namespace GraphMaker
{
    public class GraphData
    {
        public string? FileName { get; set; }
        public int XColumnIndex { get; set; }
        public List<int> YColumnIndices { get; set; }
        public int? SpecColumnIndex { get; set; }
        public int? UpperLimitColumnIndex { get; set; }
        public int? LowerLimitColumnIndex { get; set; }
        public DataTable Data { get; set; }
        public List<string> HeaderRow { get; set; }
        public Dictionary<int, int> ColumnGroupMap { get; set; } = new Dictionary<int, int>();
        public Color SpecColor { get; set; } = Colors.Black;
        public Color UpperColor { get; set; } = Colors.Black;
        public Color LowerColor { get; set; } = Colors.Black;
        public Color YAxisColor { get; set; } = Colors.Black;
        public LineStyle SpecLineStyle { get; set; } = LineStyle.Dash;
        public LineStyle UpperLineStyle { get; set; } = LineStyle.Dash;
        public LineStyle LowerLineStyle { get; set; } = LineStyle.Dash;
        public Dictionary<int, string> GroupNames { get; set; } = new Dictionary<int, string>();
        public Dictionary<int, Color> GroupColors { get; set; } = new Dictionary<int, Color>();
        public Dictionary<int, LineStyle> GroupLineStyles { get; set; } = new Dictionary<int, LineStyle>();
        public Dictionary<int, int> GroupRefColumnIndices { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> GroupUpperLimitColumnIndices { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> GroupLowerLimitColumnIndices { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, Color> GroupRefColors { get; set; } = new Dictionary<int, Color>();
        public Dictionary<int, Color> GroupUpperColors { get; set; } = new Dictionary<int, Color>();
        public Dictionary<int, Color> GroupLowerColors { get; set; } = new Dictionary<int, Color>();
        public Dictionary<int, LineStyle> GroupRefLineStyles { get; set; } = new Dictionary<int, LineStyle>();
        public Dictionary<int, LineStyle> GroupUpperLineStyles { get; set; } = new Dictionary<int, LineStyle>();
        public Dictionary<int, LineStyle> GroupLowerLineStyles { get; set; } = new Dictionary<int, LineStyle>();
    }

    public partial class GraphViewerWindow : Window, INotifyPropertyChanged
    {
        private PlotModel _lineChartModel;
        private PlotModel _scatterChartModel;
        private PlotModel _averageChartModel;
        private PlotModel _cpkChartModel;
        private readonly bool _useLogXAxis;

        public PlotModel LineChartModel
        {
            get => _lineChartModel;
            set { _lineChartModel = value; OnPropertyChanged(nameof(LineChartModel)); }
        }

        public PlotModel ScatterChartModel
        {
            get => _scatterChartModel;
            set { _scatterChartModel = value; OnPropertyChanged(nameof(ScatterChartModel)); }
        }

        public PlotModel AverageChartModel
        {
            get => _averageChartModel;
            set { _averageChartModel = value; OnPropertyChanged(nameof(AverageChartModel)); }
        }

        public PlotModel CpkChartModel
        {
            get => _cpkChartModel;
            set { _cpkChartModel = value; OnPropertyChanged(nameof(CpkChartModel)); }
        }

        private readonly List<GraphData> _graphDataList;

        public GraphViewerWindow(List<GraphData> graphDataList, bool useLogXAxis = false)
        {
            InitializeComponent();
            DataContext = this;
            _graphDataList = graphDataList;
            if (useLogXAxis && ContainsNonPositiveXValues(graphDataList))
            {
                MessageBox.Show("X axis has values <= 0. Log scale is disabled for this graph.",
                    "X Axis Log Scale", MessageBoxButton.OK, MessageBoxImage.Information);
                _useLogXAxis = false;
            }
            else
            {
                _useLogXAxis = useLogXAxis;
            }

            GenerateCharts();
        }

        private void GenerateCharts()
        {
            try
            {
                LineChartModel = CreateLineChart(_useLogXAxis);
                ScatterChartModel = CreateNormalDistributionChart();
                AverageChartModel = CreateAverageChart(_useLogXAxis);
                CpkChartModel = CreateCpkChart(_useLogXAxis);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to create graph:\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LineHeader_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenSingleGraphWindow("Line Graph", CreateLineChart(_useLogXAxis));
        }

        private void ScatterHeader_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenSingleGraphWindow("Normal Distribution", CreateNormalDistributionChart());
        }

        private void AverageHeader_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenSingleGraphWindow("Average Graph", CreateAverageChart(_useLogXAxis));
        }

        private void CpkHeader_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            OpenSingleGraphWindow("CPK Graph", CreateCpkChart(_useLogXAxis));
        }

        private void OpenSingleGraphWindow(string title, PlotModel model)
        {
            LargePlotWindowHelper.CreateLargePlotWindow(this, title, model).Show();
        }

        private PlotModel CreateLineChart(bool useLogXAxis)
        {
            var model = new PlotModel { Title = "Line Plot" };
            ApplyCommonPlotStyle(model);
            model.Axes.Add(CreateBottomAxis("X", useLogXAxis));
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Y",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            var drawnLimitKeys = new HashSet<string>();

            foreach (var graphData in _graphDataList)
            {
                var limitKey = GetLimitKey(graphData);
                if (!string.IsNullOrEmpty(limitKey) && drawnLimitKeys.Add(limitKey))
                {
                    AddLimitLines(model, graphData);
                }

                foreach (var yIndex in graphData.YColumnIndices)
                {
                    if (IsLimitColumn(graphData, yIndex))
                    {
                        continue;
                    }

                    int groupNumber = GetGroupNumber(graphData, yIndex);
                    var groupName = GetGroupName(graphData, groupNumber);
                    var seriesName = groupNumber > 0
                        ? $"{graphData.HeaderRow[yIndex]} ({groupName})"
                        : $"{graphData.HeaderRow[yIndex]}";

                    var lineSeries = new LineSeries
                    {
                        Title = seriesName,
                        RenderInLegend = true,
                        Color = ToOxyColor(GetSeriesColor(graphData, groupNumber)),
                        StrokeThickness = 2,
                        LineStyle = GetSeriesLineStyle(graphData, groupNumber)
                    };

                    int rowIndex = 0;
                    foreach (DataRow row in graphData.Data.Rows)
                    {
                        if (TryGetXValue(graphData, row, rowIndex, out double x) &&
                            double.TryParse(row[yIndex]?.ToString(), out double y))
                        {
                            lineSeries.Points.Add(new DataPoint(x, y));
                        }
                        rowIndex++;
                    }

                    model.Series.Add(lineSeries);
                }
            }

            model.IsLegendVisible = true;
            return model;
        }

        private PlotModel CreateScatterChart(bool useLogXAxis)
        {
            var model = new PlotModel { Title = "Scatter by X (Group Min/Avg/Max)" };
            ApplyCommonPlotStyle(model);
            model.Axes.Add(CreateBottomAxis("X", useLogXAxis));
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Y",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            var drawnLimitKeys = new HashSet<string>();
            int totalGroups = 0;
            int ngGroups = 0;

            foreach (var graphData in _graphDataList)
            {
                var limitKey = GetLimitKey(graphData);
                if (!string.IsNullOrEmpty(limitKey) && drawnLimitKeys.Add(limitKey))
                {
                    AddLimitLines(model, graphData);
                }

                foreach (var grouped in GetYAxisColumnsByGroup(graphData))
                {
                    int groupNumber = grouped.Key;
                    var yColumns = grouped.Value;
                    if (yColumns.Count == 0)
                    {
                        continue;
                    }

                    var groupName = GetGroupName(graphData, groupNumber);

                    bool isNg = false;
                    var hasLimitForSeries = TryGetGroupLimitColumns(graphData, groupNumber, out _, out int upperCol, out int lowerCol);
                    var upperLimits = new Dictionary<double, double>();
                    var lowerLimits = new Dictionary<double, double>();

                    if (hasLimitForSeries)
                    {
                        int limitRow = 0;
                        foreach (DataRow row in graphData.Data.Rows)
                        {
                            if (TryGetXValue(graphData, row, limitRow, out double x) &&
                                double.TryParse(row[upperCol]?.ToString(), out double upper) &&
                                double.TryParse(row[lowerCol]?.ToString(), out double lower))
                            {
                                upperLimits[x] = upper;
                                lowerLimits[x] = lower;
                            }
                            limitRow++;
                        }
                    }

                    var xValues = new Dictionary<double, List<double>>();
                    int rowIndex = 0;
                    foreach (DataRow row in graphData.Data.Rows)
                    {
                        if (!TryGetXValue(graphData, row, rowIndex, out double x))
                        {
                            rowIndex++;
                            continue;
                        }

                        foreach (var yIndex in yColumns)
                        {
                            if (!double.TryParse(row[yIndex]?.ToString(), out double y))
                            {
                                continue;
                            }

                            if (!xValues.TryGetValue(x, out var values))
                            {
                                values = new List<double>();
                                xValues[x] = values;
                            }

                            values.Add(y);
                            if (hasLimitForSeries && upperLimits.TryGetValue(x, out double upper) && lowerLimits.TryGetValue(x, out double lower))
                            {
                                if (y > upper || y < lower) isNg = true;
                            }
                        }
                        rowIndex++;
                    }

                    var baseColor = ToOxyColor(GetSeriesColor(graphData, groupNumber));
                    var prefix = groupName;
                    var minSeries = new LineSeries
                    {
                        Title = $"{prefix}_min",
                        RenderInLegend = true,
                        Color = OxyColor.FromAColor(190, baseColor),
                        StrokeThickness = 1.5,
                        LineStyle = LineStyle.Dot
                    };
                    var avgSeries = new LineSeries
                    {
                        Title = $"{prefix}_avg",
                        RenderInLegend = true,
                        Color = baseColor,
                        StrokeThickness = 2.4,
                        LineStyle = LineStyle.Solid
                    };
                    var maxSeries = new LineSeries
                    {
                        Title = $"{prefix}_max",
                        RenderInLegend = true,
                        Color = OxyColor.FromAColor(190, baseColor),
                        StrokeThickness = 1.5,
                        LineStyle = LineStyle.Dash
                    };

                    foreach (var pair in xValues.OrderBy(p => p.Key))
                    {
                        if (pair.Value.Count == 0) continue;
                        minSeries.Points.Add(new DataPoint(pair.Key, pair.Value.Min()));
                        avgSeries.Points.Add(new DataPoint(pair.Key, pair.Value.Average()));
                        maxSeries.Points.Add(new DataPoint(pair.Key, pair.Value.Max()));
                    }

                    model.Series.Add(minSeries);
                    model.Series.Add(avgSeries);
                    model.Series.Add(maxSeries);
                    totalGroups++;
                    if (isNg) ngGroups++;
                }
            }

            if (totalGroups > 0)
            {
                double ngRate = (double)ngGroups / totalGroups * 100.0;
                model.Title = $"Plot of Group NG Groups: {ngGroups}/{totalGroups} ({ngRate:F1}%)";
            }

            model.IsLegendVisible = true;
            return model;
        }

        private PlotModel CreateNormalDistributionChart()
        {
            var model = new PlotModel { Title = "Normal Distribution" };
            ApplyCommonPlotStyle(model);
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Value",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Density",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            foreach (var graphData in _graphDataList)
            {
                foreach (var grouped in GetYAxisColumnsByGroup(graphData))
                {
                    int groupNumber = grouped.Key;
                    var yColumns = grouped.Value;
                    if (yColumns.Count == 0)
                    {
                        continue;
                    }

                    var values = new List<double>();
                    foreach (DataRow row in graphData.Data.Rows)
                    {
                        foreach (var yIndex in yColumns)
                        {
                            if (double.TryParse(row[yIndex]?.ToString(), out double y))
                            {
                                values.Add(y);
                            }
                        }
                    }

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

                    double min = values.Min();
                    double max = values.Max();
                    if (Math.Abs(max - min) <= 1e-12)
                    {
                        min = mean - (3 * stdDev);
                        max = mean + (3 * stdDev);
                    }
                    else
                    {
                        min -= stdDev;
                        max += stdDev;
                    }

                    string groupName = GetGroupName(graphData, groupNumber);
                    string seriesTitle = string.IsNullOrWhiteSpace(groupName) ? "Distribution" : $"{groupName}_dist";
                    var series = new LineSeries
                    {
                        Title = seriesTitle,
                        RenderInLegend = true,
                        Color = ToOxyColor(GetSeriesColor(graphData, groupNumber)),
                        StrokeThickness = 2,
                        LineStyle = GetSeriesLineStyle(graphData, groupNumber)
                    };

                    const int segments = 100;
                    for (int i = 0; i <= segments; i++)
                    {
                        double x = min + ((max - min) * i / segments);
                        double y = (1.0 / (stdDev * Math.Sqrt(2 * Math.PI))) *
                                   Math.Exp(-0.5 * Math.Pow((x - mean) / stdDev, 2));
                        series.Points.Add(new DataPoint(x, y));
                    }

                    model.Series.Add(series);
                }
            }

            model.IsLegendVisible = true;
            return model;
        }

        private PlotModel CreateAverageChart(bool useLogXAxis)
        {
            var model = new PlotModel { Title = "Plot of Average Each Group" };
            ApplyCommonPlotStyle(model);
            model.Axes.Add(CreateBottomAxis("X", useLogXAxis));
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Average",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            var drawnLimitKeys = new HashSet<string>();

            foreach (var graphData in _graphDataList)
            {
                var limitKey = GetLimitKey(graphData);
                if (!string.IsNullOrEmpty(limitKey) && drawnLimitKeys.Add(limitKey))
                {
                    AddLimitLines(model, graphData);
                }

                foreach (var grouped in GetYAxisColumnsByGroup(graphData))
                {
                    int groupNumber = grouped.Key;
                    var yColumns = grouped.Value;
                    if (yColumns.Count == 0)
                    {
                        continue;
                    }

                    var xGroups = new Dictionary<double, List<double>>();
                    int rowIndex = 0;
                    foreach (DataRow row in graphData.Data.Rows)
                    {
                        if (!TryGetXValue(graphData, row, rowIndex, out double x))
                        {
                            rowIndex++;
                            continue;
                        }

                        foreach (var yIndex in yColumns)
                        {
                            if (!double.TryParse(row[yIndex]?.ToString(), out double y))
                            {
                                continue;
                            }

                            if (!xGroups.TryGetValue(x, out var bucket))
                            {
                                bucket = new List<double>();
                                xGroups[x] = bucket;
                            }
                            bucket.Add(y);
                        }
                        rowIndex++;
                    }

                    var groupName = GetGroupName(graphData, groupNumber);
                    var baseColor = ToOxyColor(GetSeriesColor(graphData, groupNumber));
                    var avgSeries = new LineSeries
                    {
                        Title = $"{groupName}_avg",
                        RenderInLegend = true,
                        Color = baseColor,
                        StrokeThickness = 2
                    };
                    var stdUpperSeries = new LineSeries
                    {
                        Title = $"{groupName}_std_plus",
                        RenderInLegend = true,
                        Color = OxyColor.FromAColor(150, baseColor),
                        StrokeThickness = 1.3,
                        LineStyle = LineStyle.Dash
                    };
                    var stdLowerSeries = new LineSeries
                    {
                        Title = $"{groupName}_std_minus",
                        RenderInLegend = true,
                        Color = OxyColor.FromAColor(150, baseColor),
                        StrokeThickness = 1.3,
                        LineStyle = LineStyle.Dash
                    };

                    foreach (var pair in xGroups.OrderBy(p => p.Key))
                    {
                        if (pair.Value.Count == 0) continue;
                        double avg = pair.Value.Average();
                        double variance = pair.Value.Sum(v => Math.Pow(v - avg, 2)) / pair.Value.Count;
                        double std = Math.Sqrt(variance);
                        avgSeries.Points.Add(new DataPoint(pair.Key, avg));
                        stdUpperSeries.Points.Add(new DataPoint(pair.Key, avg + std));
                        stdLowerSeries.Points.Add(new DataPoint(pair.Key, avg - std));
                    }

                    model.Series.Add(avgSeries);
                    model.Series.Add(stdUpperSeries);
                    model.Series.Add(stdLowerSeries);
                }
            }

            model.IsLegendVisible = true;
            return model;
        }

        private PlotModel CreateCpkChart(bool useLogXAxis)
        {
            var model = new PlotModel { Title = "X-CPK" };
            ApplyCommonPlotStyle(model);
            model.Axes.Add(CreateBottomAxis("X", useLogXAxis));
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "CPK",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });

            foreach (var graphData in _graphDataList)
            {
                foreach (var grouped in GetYAxisColumnsByGroup(graphData))
                {
                    int groupNumber = grouped.Key;
                    var yColumns = grouped.Value;
                    if (yColumns.Count == 0)
                    {
                        continue;
                    }

                    if (!TryGetGroupLimitColumns(graphData, groupNumber, out _, out int upperCol, out int lowerCol))
                    {
                        if (!TryGetSeriesLimitColumns(graphData, yColumns[0], out _, out upperCol, out lowerCol))
                        {
                            continue;
                        }
                    }

                    var xGroups = new Dictionary<double, List<double>>();
                    var xLimitGroups = new Dictionary<double, List<(double Upper, double Lower)>>();
                    int rowIndex = 0;
                    foreach (DataRow row in graphData.Data.Rows)
                    {
                        if (!TryGetXValue(graphData, row, rowIndex, out double x))
                        {
                            rowIndex++;
                            continue;
                        }

                        if (double.TryParse(row[upperCol]?.ToString(), out double upper) &&
                            double.TryParse(row[lowerCol]?.ToString(), out double lower))
                        {
                            if (!xLimitGroups.TryGetValue(x, out var limits))
                            {
                                limits = new List<(double Upper, double Lower)>();
                                xLimitGroups[x] = limits;
                            }
                            limits.Add((upper, lower));
                        }

                        foreach (var yIndex in yColumns)
                        {
                            if (!double.TryParse(row[yIndex]?.ToString(), out double y))
                            {
                                continue;
                            }

                            if (!xGroups.TryGetValue(x, out var values))
                            {
                                values = new List<double>();
                                xGroups[x] = values;
                            }

                            values.Add(y);
                        }
                        rowIndex++;
                    }

                    var groupName = GetGroupName(graphData, groupNumber);
                    var series = new LineSeries
                    {
                        Title = $"{groupName}_cpk",
                        RenderInLegend = true,
                        Color = ToOxyColor(GetSeriesColor(graphData, groupNumber)),
                        StrokeThickness = 2,
                        LineStyle = GetSeriesLineStyle(graphData, groupNumber)
                    };

                    foreach (var pair in xGroups.OrderBy(p => p.Key))
                    {
                        if (!xLimitGroups.TryGetValue(pair.Key, out var limitsAtX) || limitsAtX.Count == 0)
                        {
                            continue;
                        }

                        double usl = limitsAtX.Average(v => v.Upper);
                        double lsl = limitsAtX.Average(v => v.Lower);
                        double mean = pair.Value.Average();
                        double variance = pair.Value.Sum(v => Math.Pow(v - mean, 2)) / pair.Value.Count;
                        double stdDev = Math.Sqrt(variance);
                        if (stdDev <= 0)
                        {
                            continue;
                        }

                        double cpkUpper = (usl - mean) / (3 * stdDev);
                        double cpkLower = (mean - lsl) / (3 * stdDev);
                        series.Points.Add(new DataPoint(pair.Key, Math.Min(cpkUpper, cpkLower)));
                    }

                    model.Series.Add(series);
                }
            }

            model.IsLegendVisible = true;
            return model;
        }

        private static void ApplyCommonPlotStyle(PlotModel model)
        {
            model.Background = OxyColors.White;
            model.PlotAreaBackground = OxyColors.White;
            model.TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.TitleColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.PlotAreaBorderColor = OxyColor.FromRgb(0xB8, 0xC7, 0xDB);
            if (model.Legends.Count == 0)
            {
                model.Legends.Add(new Legend
                {
                    LegendPlacement = LegendPlacement.Outside,
                    LegendPosition = LegendPosition.RightTop,
                    LegendOrientation = LegendOrientation.Vertical,
                    LegendBackground = OxyColor.FromAColor(245, OxyColors.White),
                    LegendBorder = OxyColor.FromRgb(0xB8, 0xC7, 0xDB),
                    TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33)
                });
            }
        }

        private string GetLimitKey(GraphData graphData)
        {
            var signatures = new List<string>();
            foreach (int groupId in GetConfiguredGroupIds(graphData))
            {
                if (!TryGetGroupLimitColumns(graphData, groupId, out int specCol, out int upperCol, out int lowerCol))
                {
                    continue;
                }

                if (TryGetLimitSignature(graphData, specCol, upperCol, lowerCol, out string signature))
                {
                    signatures.Add($"G{groupId}:{signature}");
                }
            }

            if (signatures.Count > 0)
            {
                return string.Join("|", signatures.OrderBy(s => s));
            }

            if (graphData.SpecColumnIndex.HasValue && graphData.UpperLimitColumnIndex.HasValue && graphData.LowerLimitColumnIndex.HasValue)
            {
                if (TryGetLimitSignature(graphData, graphData.SpecColumnIndex.Value, graphData.UpperLimitColumnIndex.Value, graphData.LowerLimitColumnIndex.Value, out string fallback))
                {
                    return fallback;
                }
            }

            return null;
        }

        private void AddLimitLines(PlotModel model, GraphData graphData)
        {
            bool addedAny = false;

            foreach (int groupId in GetConfiguredGroupIds(graphData))
            {
                if (!TryGetGroupLimitColumns(graphData, groupId, out int specCol, out int upperCol, out int lowerCol))
                {
                    continue;
                }

                var groupName = GetGroupName(graphData, groupId);
                var refColor = graphData.GroupRefColors.TryGetValue(groupId, out Color gRefColor) ? gRefColor : graphData.SpecColor;
                var upperColor = graphData.GroupUpperColors.TryGetValue(groupId, out Color gUpperColor) ? gUpperColor : graphData.UpperColor;
                var lowerColor = graphData.GroupLowerColors.TryGetValue(groupId, out Color gLowerColor) ? gLowerColor : graphData.LowerColor;
                var refLine = graphData.GroupRefLineStyles.TryGetValue(groupId, out LineStyle gRefLine) ? gRefLine : graphData.SpecLineStyle;
                var upperLine = graphData.GroupUpperLineStyles.TryGetValue(groupId, out LineStyle gUpperLine) ? gUpperLine : graphData.UpperLineStyle;
                var lowerLine = graphData.GroupLowerLineStyles.TryGetValue(groupId, out LineStyle gLowerLine) ? gLowerLine : graphData.LowerLineStyle;

                AddLimitLinesForColumns(model, graphData, specCol, upperCol, lowerCol,
                    $"{groupName} REF", $"{groupName} Upper", $"{groupName} Lower",
                    refColor, upperColor, lowerColor, refLine, upperLine, lowerLine);
                addedAny = true;
            }

            if (!addedAny && graphData.SpecColumnIndex.HasValue && graphData.UpperLimitColumnIndex.HasValue && graphData.LowerLimitColumnIndex.HasValue)
            {
                AddLimitLinesForColumns(model, graphData,
                    graphData.SpecColumnIndex.Value,
                    graphData.UpperLimitColumnIndex.Value,
                    graphData.LowerLimitColumnIndex.Value,
                    "REF", "Upper Limit", "Lower Limit",
                    graphData.SpecColor, graphData.UpperColor, graphData.LowerColor,
                    graphData.SpecLineStyle, graphData.UpperLineStyle, graphData.LowerLineStyle);
            }
        }

        private static void AddLimitLinesForColumns(
            PlotModel model,
            GraphData graphData,
            int specCol,
            int upperCol,
            int lowerCol,
            string refTitle,
            string upperTitle,
            string lowerTitle,
            Color refColor,
            Color upperColor,
            Color lowerColor,
            LineStyle refLineStyle,
            LineStyle upperLineStyle,
            LineStyle lowerLineStyle)
        {
            var specPoints = new List<DataPoint>();
            var upperPoints = new List<DataPoint>();
            var lowerPoints = new List<DataPoint>();

            int rowIndex = 0;
            foreach (DataRow row in graphData.Data.Rows)
            {
                if (TryGetXValue(graphData, row, rowIndex, out double x) &&
                    double.TryParse(row[specCol]?.ToString(), out double spec) &&
                    double.TryParse(row[upperCol]?.ToString(), out double upper) &&
                    double.TryParse(row[lowerCol]?.ToString(), out double lower))
                {
                    specPoints.Add(new DataPoint(x, spec));
                    upperPoints.Add(new DataPoint(x, upper));
                    lowerPoints.Add(new DataPoint(x, lower));
                }
                rowIndex++;
            }

            if (specPoints.Count == 0)
            {
                return;
            }

            var specSeries = new LineSeries
            {
                Title = refTitle,
                RenderInLegend = true,
                Color = ToOxyColor(refColor),
                StrokeThickness = 2,
                LineStyle = refLineStyle
            };
            specSeries.Points.AddRange(specPoints);
            model.Series.Add(specSeries);

            var upperSeries = new LineSeries
            {
                Title = upperTitle,
                RenderInLegend = true,
                Color = ToOxyColor(upperColor),
                StrokeThickness = 2,
                LineStyle = upperLineStyle
            };
            upperSeries.Points.AddRange(upperPoints);
            model.Series.Add(upperSeries);

            var lowerSeries = new LineSeries
            {
                Title = lowerTitle,
                RenderInLegend = true,
                Color = ToOxyColor(lowerColor),
                StrokeThickness = 2,
                LineStyle = lowerLineStyle
            };
            lowerSeries.Points.AddRange(lowerPoints);
            model.Series.Add(lowerSeries);
        }

        private static bool TryGetLimitSignature(GraphData graphData, int specCol, int upperCol, int lowerCol, out string signature)
        {
            double spec = 0;
            double upper = 0;
            double lower = 0;
            int count = 0;

            foreach (DataRow row in graphData.Data.Rows)
            {
                if (double.TryParse(row[specCol]?.ToString(), out double s) &&
                    double.TryParse(row[upperCol]?.ToString(), out double u) &&
                    double.TryParse(row[lowerCol]?.ToString(), out double l))
                {
                    spec += s;
                    upper += u;
                    lower += l;
                    count++;
                }
            }

            if (count == 0)
            {
                signature = string.Empty;
                return false;
            }

            signature = $"{spec / count:F4}_{upper / count:F4}_{lower / count:F4}";
            return true;
        }

        private static bool TryGetSeriesLimitColumns(GraphData graphData, int yIndex, out int specCol, out int upperCol, out int lowerCol)
        {
            int group = GetGroupNumber(graphData, yIndex);
            if (group > 0 && TryGetGroupLimitColumns(graphData, group, out specCol, out upperCol, out lowerCol))
            {
                return true;
            }

            if (graphData.SpecColumnIndex.HasValue && graphData.UpperLimitColumnIndex.HasValue && graphData.LowerLimitColumnIndex.HasValue)
            {
                specCol = graphData.SpecColumnIndex.Value;
                upperCol = graphData.UpperLimitColumnIndex.Value;
                lowerCol = graphData.LowerLimitColumnIndex.Value;
                return true;
            }

            specCol = -1;
            upperCol = -1;
            lowerCol = -1;
            return false;
        }

        private static bool TryGetGroupLimitColumns(GraphData graphData, int groupId, out int specCol, out int upperCol, out int lowerCol)
        {
            if (graphData.GroupRefColumnIndices.TryGetValue(groupId, out specCol) &&
                graphData.GroupUpperLimitColumnIndices.TryGetValue(groupId, out upperCol) &&
                graphData.GroupLowerLimitColumnIndices.TryGetValue(groupId, out lowerCol) &&
                specCol >= 0 && upperCol >= 0 && lowerCol >= 0)
            {
                return true;
            }

            specCol = -1;
            upperCol = -1;
            lowerCol = -1;
            return false;
        }

        private static bool IsLimitColumn(GraphData graphData, int columnIndex)
        {
            if (graphData.SpecColumnIndex == columnIndex ||
                graphData.UpperLimitColumnIndex == columnIndex ||
                graphData.LowerLimitColumnIndex == columnIndex)
            {
                return true;
            }

            return graphData.GroupRefColumnIndices.Values.Contains(columnIndex)
                || graphData.GroupUpperLimitColumnIndices.Values.Contains(columnIndex)
                || graphData.GroupLowerLimitColumnIndices.Values.Contains(columnIndex);
        }

        private static IEnumerable<int> GetConfiguredGroupIds(GraphData graphData)
        {
            return graphData.GroupRefColumnIndices.Keys
                .Concat(graphData.GroupUpperLimitColumnIndices.Keys)
                .Concat(graphData.GroupLowerLimitColumnIndices.Keys)
                .Distinct();
        }

        private static bool TryGetXValue(GraphData graphData, DataRow row, int rowIndex, out double x)
        {
            if (graphData.XColumnIndex >= 0)
            {
                return double.TryParse(row[graphData.XColumnIndex]?.ToString(), out x);
            }

            x = rowIndex + 1;
            return true;
        }

        private static int GetGroupNumber(GraphData graphData, int yIndex)
        {
            return graphData.ColumnGroupMap.TryGetValue(yIndex, out int group) && group > 0 ? group : 0;
        }

        private static Dictionary<int, List<int>> GetYAxisColumnsByGroup(GraphData graphData)
        {
            var result = new Dictionary<int, List<int>>();
            foreach (var yIndex in graphData.YColumnIndices)
            {
                if (IsLimitColumn(graphData, yIndex))
                {
                    continue;
                }

                int group = GetGroupNumber(graphData, yIndex);
                if (!result.TryGetValue(group, out var list))
                {
                    list = new List<int>();
                    result[group] = list;
                }

                list.Add(yIndex);
            }

            return result;
        }

        private static string GetGroupName(GraphData graphData, int groupNumber)
        {
            if (groupNumber <= 0)
            {
                return string.Empty;
            }

            if (graphData.GroupNames.TryGetValue(groupNumber, out string? name) && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            return $"Group {groupNumber}";
        }

        private static Color GetSeriesColor(GraphData graphData, int groupNumber)
        {
            if (groupNumber > 0 && graphData.GroupColors.TryGetValue(groupNumber, out Color color))
            {
                return color;
            }

            return graphData.YAxisColor;
        }

        private static LineStyle GetSeriesLineStyle(GraphData graphData, int groupNumber)
        {
            if (groupNumber > 0 && graphData.GroupLineStyles.TryGetValue(groupNumber, out LineStyle style))
            {
                return style;
            }

            return LineStyle.Solid;
        }

        private static OxyColor ToOxyColor(Color color)
        {
            return OxyColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        private static bool ContainsNonPositiveXValues(IEnumerable<GraphData> graphDataList)
        {
            foreach (var graphData in graphDataList)
            {
                int rowIndex = 0;
                foreach (DataRow row in graphData.Data.Rows)
                {
                    if (TryGetXValue(graphData, row, rowIndex, out double x) && x <= 0)
                    {
                        return true;
                    }
                    rowIndex++;
                }
            }

            return false;
        }

        private static Axis CreateBottomAxis(string title, bool useLogXAxis)
        {
            if (useLogXAxis)
            {
                return new LogarithmicAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = title,
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot
                };
            }

            return new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Title = title,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
