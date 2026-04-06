using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Annotations;
using OxyPlot.Legends;
using OxyPlot.Wpf;
using Clipboard = System.Windows.Clipboard;
using Color = System.Windows.Media.Color;
using MessageBox = System.Windows.MessageBox;
using TextDataFormat = System.Windows.TextDataFormat;

namespace GraphMaker
{
    public class DailySamplingSummaryRow
    {
        public string XAxisSeriesName { get; set; } = string.Empty;
        public double Max { get; set; }
        public double Min { get; set; }
        public double Avg { get; set; }
        public double? Cpk { get; set; }
        public int TotalCount { get; set; }
        public int NgCount { get; set; }
        public double Ppm { get; set; }
    }

    internal sealed class SummaryValueItem
    {
        public double Value { get; init; }
        public double? Upper { get; init; }
        public double? Lower { get; init; }
    }

    public partial class DailySamplingGraphViewerWindow : Window
    {
        private readonly List<DailySamplingGraphData> _graphDataList;
        private readonly Func<string, DateTime?> _parseDate;
        private bool _isInitializingView;

        public PlotModel ScatterPlotModel { get; set; }
        public PlotModel CpkPlotModel { get; set; }
        public List<DailySamplingSummaryRow> SummaryRows { get; set; } = new List<DailySamplingSummaryRow>();

        public DailySamplingGraphViewerWindow(List<DailySamplingGraphData> graphDataList, ValuePlotView parentView)
            : this(graphDataList, parentView.ParseDateString)
        {
        }

        public DailySamplingGraphViewerWindow(List<DailySamplingGraphData> graphDataList, ValuePlotMultiColumnView parentView)
            : this(graphDataList, parentView.ParseDateString)
        {
        }

        public DailySamplingGraphViewerWindow(List<DailySamplingGraphData> graphDataList, DailyDataTrendView parentView)
            : this(graphDataList, parentView.ParseDateString)
        {
        }

        public DailySamplingGraphViewerWindow(List<DailySamplingGraphData> graphDataList, UnifiedMultiYView parentView)
            : this(graphDataList, parentView.ParseDateString)
        {
        }

        private DailySamplingGraphViewerWindow(List<DailySamplingGraphData> graphDataList, Func<string, DateTime?> parseDate)
        {
            InitializeComponent();
            _graphDataList = graphDataList?.Where(item => item != null).ToList() ?? new List<DailySamplingGraphData>();
            _parseDate = parseDate;

            DataContext = this;
            InitializeSummaryPeriod();
            RefreshView();
        }

        private void ScatterHeader_Click(object sender, RoutedEventArgs e)
        {
            OpenSingleGraphWindow("Daily Sampling Scatter Plot", BuildScatterPlotModel());
        }

        private void CpkHeader_Click(object sender, RoutedEventArgs e)
        {
            OpenSingleGraphWindow("CPK by Date", BuildCpkPlotModel());
        }

        private void OpenSingleGraphWindow(string title, PlotModel model)
        {
            LargePlotWindowHelper.CreateLargePlotWindow(this, title, model).Show();
        }

        private void CreateScatterPlot()
        {
            ScatterPlotModel = BuildScatterPlotModel();
        }

        private void InitializeSummaryPeriod()
        {
            _isInitializingView = true;
            try
            {
                var dates = _graphDataList
                    .SelectMany(GetParsedDates)
                    .OrderBy(date => date)
                    .ToList();

                if (dates.Count == 0)
                {
                    SummaryFromDatePicker.IsEnabled = false;
                    SummaryToDatePicker.IsEnabled = false;
                    return;
                }

                SummaryFromDatePicker.SelectedDate = dates.First().Date;
                SummaryToDatePicker.SelectedDate = dates.Last().Date;
            }
            finally
            {
                _isInitializingView = false;
            }
        }

        private void RefreshView()
        {
            CreateScatterPlot();
            CreateCpkPlot();
            BuildSummaryRows();

            ScatterPlotView.Model = ScatterPlotModel;
            ScatterPlotView.InvalidatePlot(true);
            CpkPlotView.Model = CpkPlotModel;
            CpkPlotView.InvalidatePlot(true);
            SummaryDataGrid.ItemsSource = null;
            SummaryDataGrid.ItemsSource = SummaryRows;
        }

        private bool ShowSampleSeries => ShowSampleCheckBox.IsChecked != false;
        private bool ShowAvgSeries => ShowAvgCheckBox.IsChecked != false;
        private bool ShowMaxAnnotations => ShowMaxCheckBox.IsChecked != false;
        private bool ShowMinAnnotations => ShowMinCheckBox.IsChecked != false;
        private bool ShowSpecSeries => ShowSpecCheckBox.IsChecked != false;
        private bool ShowUpperLimitSeries => ShowUpperLimitCheckBox.IsChecked != false;
        private bool ShowLowerLimitSeries => ShowLowerLimitCheckBox.IsChecked != false;
        private bool ShowCpkSeries => ShowCpkCheckBox.IsChecked != false;
        private DateTime? SummaryFromDate => SummaryFromDatePicker.SelectedDate?.Date;
        private DateTime? SummaryToDate => SummaryToDatePicker.SelectedDate?.Date;

        private IEnumerable<DateTime> GetParsedDates(DailySamplingGraphData graphData)
        {
            if (graphData.Dates == null)
            {
                yield break;
            }

            foreach (string dateText in graphData.Dates)
            {
                DateTime? parsed = _parseDate(dateText);
                if (parsed.HasValue)
                {
                    yield return parsed.Value.Date;
                }
            }
        }

        private PlotModel BuildScatterPlotModel()
        {
            var model = new PlotModel { Title = "Daily Sampling Scatter Plot" };
            ApplyCommonPlotStyle(model);

            // X축 타입 결정 (첫 번째 파일 기준)
            DailySamplingGraphData? firstGraphData = _graphDataList.FirstOrDefault();
            var xAxisMode = firstGraphData?.XAxisMode ?? XAxisMode.Date;
            bool useDate = xAxisMode == XAxisMode.Date;
            bool useFixedSingleX = xAxisMode == XAxisMode.None &&
                                   firstGraphData?.NoXAxisDisplayMode == NoXAxisDisplayMode.FixedSingleX;
            bool splitByColumn = firstGraphData?.DisplayMode == ValuePlotDisplayMode.SplitByColumn;

            // 전체 NG Rate 계산용 변수
            int totalNgPoints = 0;
            int totalPoints = 0;

            // X축 설정
            if (useDate)
            {
                // X축: 날짜 (DateTimeAxis)
                var dateAxis = new DateTimeAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Date",
                    StringFormat = "MM-dd",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IntervalType = DateTimeIntervalType.Days
                };
                model.Axes.Add(dateAxis);
            }
            else
            {
                // X축: 순번 (LinearAxis)
                var sequenceAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = useFixedSingleX ? string.Empty :
                        (xAxisMode == XAxisMode.None ? "Sample Order" : "Sequence"),
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IsAxisVisible = !useFixedSingleX
                };
                model.Axes.Add(sequenceAxis);
            }

            // Y축: 샘플 값
            var valueAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Value",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };
            model.Axes.Add(valueAxis);

            // 각 파일에 대해 산포도 생성
            foreach (var graphData in _graphDataList)
            {
                if (splitByColumn)
                {
                    AddScatterSeriesByColumn(model, graphData, useDate, useFixedSingleX, ref totalNgPoints, ref totalPoints);
                    continue;
                }

                var scatterSeries = new ScatterSeries
                {
                    Title = "Sample",
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                    graphData.DataColor.G, graphData.DataColor.B)
                };

                if (useDate)
                {
                    // 날짜별 통계를 위한 딕셔너리 (중복 날짜 처리)
                    var dateStats = new Dictionary<DateTime, List<double>>();
                    var dateNgCounts = new Dictionary<DateTime, int>();  // 날짜별 NG 포인트 수
                    var dateTotalCounts = new Dictionary<DateTime, int>();  // 날짜별 총 포인트 수

                    // 각 행의 날짜를 확인하여 중복 처리
                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        var dateStr = graphData.Dates[rowIndex];
                        var parsedDate = _parseDate(dateStr);

                        if (parsedDate.HasValue)
                        {
                            var dateValue = DateTimeAxis.ToDouble(parsedDate.Value);

                            // 날짜별로 값 모으기 (중복 날짜의 경우 같은 리스트에 추가됨)
                            if (!dateStats.ContainsKey(parsedDate.Value))
                            {
                                dateStats[parsedDate.Value] = new List<double>();
                                dateNgCounts[parsedDate.Value] = 0;
                                dateTotalCounts[parsedDate.Value] = 0;
                            }

                            // 해당 날짜의 모든 샘플 값 (2열부터 끝까지)
                            var dataRow = graphData.Data.Rows[rowIndex];
                            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                            {
                                var valueStr = dataRow[colIndex]?.ToString();
                                if (!string.IsNullOrWhiteSpace(valueStr) &&
                                    double.TryParse(valueStr, out double value))
                                {
                                    scatterSeries.Points.Add(new ScatterPoint(dateValue, value));
                                    dateStats[parsedDate.Value].Add(value);

                                    // NG 체크 (Upper/Lower Limit이 있는 경우)
                                    var rowLimits = GetLimitsForColumn(graphData, colIndex);
                                    if (rowLimits.Upper.HasValue && rowLimits.Lower.HasValue)
                                    {
                                        dateTotalCounts[parsedDate.Value]++;
                                        totalPoints++;

                                        if (value > rowLimits.Upper.Value || value < rowLimits.Lower.Value)
                                        {
                                            dateNgCounts[parsedDate.Value]++;
                                            totalNgPoints++;
                                        }
                                    }
                                }
                            }
                        }
                    }

                        model.Series.Add(scatterSeries);

                    // Avg 선 추가 (값은 표시하지 않음, 선만)
                    var avgSeries = new ScatterSeries
                    {
                        Title = "Avg",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 5,
                        MarkerFill = OxyColors.Green
                    };
                    var maxSeries = new ScatterSeries
                    {
                        Title = "Max",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        MarkerFill = OxyColors.Red
                    };
                    var minSeries = new ScatterSeries
                    {
                        Title = "Min",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        MarkerFill = OxyColors.Blue
                    };

                    // 날짜별 통계 계산 및 텍스트 주석 추가
                    foreach (var kvp in dateStats.OrderBy(x => x.Key))
                    {
                        var date = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var dateValue = DateTimeAxis.ToDouble(date);
                            var max = values.Max();
                            var min = values.Min();
                            var avg = values.Average();

                            // Avg 선에 포인트 추가
                            avgSeries.Points.Add(new ScatterPoint(dateValue, avg));
                            maxSeries.Points.Add(new ScatterPoint(dateValue, max));
                            minSeries.Points.Add(new ScatterPoint(dateValue, min));

                            // Max 텍스트 주석
                            if (ShowMaxAnnotations)
                            {
                                var maxAnnotation = new TextAnnotation
                                {
                                    Text = $"Max: {max:F2}",
                                    TextPosition = new DataPoint(dateValue, max),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColors.Red
                                };
                                model.Annotations.Add(maxAnnotation);
                            }

                            // Min 텍스트 주석
                            if (ShowMinAnnotations)
                            {
                                var minAnnotation = new TextAnnotation
                                {
                                    Text = $"Min: {min:F2}",
                                    TextPosition = new DataPoint(dateValue, min),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Top,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColors.Blue
                                };
                                model.Annotations.Add(minAnnotation);
                            }

                            // Avg 텍스트 주석
                            var avgAnnotation = new TextAnnotation
                            {
                                Text = $"Avg: {avg:F2}",
                                TextPosition = new DataPoint(dateValue, avg),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Green
                            };
                            model.Annotations.Add(avgAnnotation);

                            // 날짜별 NG Rate 주석 추가 (Limit이 있는 경우)
                            if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue &&
                                dateTotalCounts.ContainsKey(date) && dateTotalCounts[date] > 0)
                            {
                                int ngCount = dateNgCounts[date];
                                int totalCount = dateTotalCounts[date];
                                double ngRate = (double)ngCount / totalCount * 100.0;

                                // Y축 중간 위치에 NG Rate 표시
                                var yPosition = (max + min) / 2.0;
                                var ngRateAnnotation = new TextAnnotation
                                {
                                    Text = $"NG: {ngCount}/{totalCount} ({ngRate:F1}%)",
                                    TextPosition = new DataPoint(dateValue, yPosition),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Right,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    FontWeight = OxyPlot.FontWeights.Bold,
                                    TextColor = ngCount > 0 ? OxyColors.Red : OxyColors.DarkGreen
                                };
                                model.Annotations.Add(ngRateAnnotation);
                            }
                        }
                    }

                    model.Series.Add(maxSeries);
                    model.Series.Add(minSeries);
                    model.Series.Add(avgSeries);
                }
                else
                {
                    // 순번별 통계를 위한 딕셔너리
                    var sequenceStats = new Dictionary<int, List<double>>();
                    var sequenceNgCounts = new Dictionary<int, int>();  // 순번별 NG 포인트 수
                    var sequenceTotalCounts = new Dictionary<int, int>();  // 순번별 총 포인트 수

                    // 각 행의 순번으로 데이터 처리
                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        int sequence = useFixedSingleX ? 1 : rowIndex + 1; // 1부터 시작하는 순번

                        // 순번별로 값 모으기
                        if (!sequenceStats.ContainsKey(sequence))
                        {
                            sequenceStats[sequence] = new List<double>();
                            sequenceNgCounts[sequence] = 0;
                            sequenceTotalCounts[sequence] = 0;
                        }

                        // 해당 순번의 모든 샘플 값 (2열부터 끝까지)
                        var dataRow = graphData.Data.Rows[rowIndex];
                        for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                        {
                            var valueStr = dataRow[colIndex]?.ToString();
                            if (!string.IsNullOrWhiteSpace(valueStr) &&
                                double.TryParse(valueStr, out double value))
                            {
                                scatterSeries.Points.Add(new ScatterPoint(sequence, value));
                                sequenceStats[sequence].Add(value);

                                // NG 체크 (Upper/Lower Limit이 있는 경우)
                                var rowLimits = GetLimitsForColumn(graphData, colIndex);
                                if (rowLimits.Upper.HasValue && rowLimits.Lower.HasValue)
                                {
                                    sequenceTotalCounts[sequence]++;
                                    totalPoints++;

                                    if (value > rowLimits.Upper.Value || value < rowLimits.Lower.Value)
                                    {
                                        sequenceNgCounts[sequence]++;
                                        totalNgPoints++;
                                    }
                                }
                            }
                        }
                    }

                    model.Series.Add(scatterSeries);

                    // Avg 선 추가
                    var avgSeries = new ScatterSeries
                    {
                        Title = "Avg",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 5,
                        MarkerFill = OxyColors.Green
                    };
                    var maxSeries = new ScatterSeries
                    {
                        Title = "Max",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        MarkerFill = OxyColors.Red
                    };
                    var minSeries = new ScatterSeries
                    {
                        Title = "Min",
                        MarkerType = MarkerType.Circle,
                        MarkerSize = 4,
                        MarkerFill = OxyColors.Blue
                    };

                    // 순번별 통계 계산 및 텍스트 주석 추가
                    foreach (var kvp in sequenceStats.OrderBy(x => x.Key))
                    {
                        var sequence = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var max = values.Max();
                            var min = values.Min();
                            var avg = values.Average();

                            // Avg 선에 포인트 추가
                            avgSeries.Points.Add(new ScatterPoint(sequence, avg));
                            maxSeries.Points.Add(new ScatterPoint(sequence, max));
                            minSeries.Points.Add(new ScatterPoint(sequence, min));

                            // Max 텍스트 주석
                            if (ShowMaxAnnotations)
                            {
                                var maxAnnotation = new TextAnnotation
                                {
                                    Text = $"Max: {max:F2}",
                                    TextPosition = new DataPoint(sequence, max),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColors.Red
                                };
                                model.Annotations.Add(maxAnnotation);
                            }

                            // Min 텍스트 주석
                            if (ShowMinAnnotations)
                            {
                                var minAnnotation = new TextAnnotation
                                {
                                    Text = $"Min: {min:F2}",
                                    TextPosition = new DataPoint(sequence, min),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Top,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColors.Blue
                                };
                                model.Annotations.Add(minAnnotation);
                            }

                            // Avg 텍스트 주석
                            var avgAnnotation = new TextAnnotation
                            {
                                Text = $"Avg: {avg:F2}",
                                TextPosition = new DataPoint(sequence, avg),
                                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                Stroke = OxyColors.Transparent,
                                FontSize = 9,
                                TextColor = OxyColors.Green
                            };
                            model.Annotations.Add(avgAnnotation);

                            // 순번별 NG Rate 주석 추가 (Limit이 있는 경우)
                            if (graphData.UpperLimitValue.HasValue && graphData.LowerLimitValue.HasValue &&
                                sequenceTotalCounts.ContainsKey(sequence) && sequenceTotalCounts[sequence] > 0)
                            {
                                int ngCount = sequenceNgCounts[sequence];
                                int totalCount = sequenceTotalCounts[sequence];
                                double ngRate = (double)ngCount / totalCount * 100.0;

                                // Y축 중간 위치에 NG Rate 표시
                                var yPosition = (max + min) / 2.0;
                                var ngRateAnnotation = new TextAnnotation
                                {
                                    Text = $"NG: {ngCount}/{totalCount} ({ngRate:F1}%)",
                                    TextPosition = new DataPoint(sequence, yPosition),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Right,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    FontWeight = OxyPlot.FontWeights.Bold,
                                    TextColor = ngCount > 0 ? OxyColors.Red : OxyColors.DarkGreen
                                };
                                model.Annotations.Add(ngRateAnnotation);
                            }
                        }
                    }

                    model.Series.Add(maxSeries);
                    model.Series.Add(minSeries);
                    model.Series.Add(avgSeries);
                }

                // SPEC/Limit 선 추가 (선택 사항)
                if (useDate)
                {
                    // 날짜 기반 SPEC/Limit 선
                    if (graphData.SpecValue.HasValue)
                    {
                        var specLine = new LineSeries
                        {
                            Title = "SPEC",
                            Color = OxyColor.FromArgb(graphData.SpecColor.A, graphData.SpecColor.R,
                                                      graphData.SpecColor.G, graphData.SpecColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        // X축 범위 계산
                        double minX = double.MaxValue;
                        double maxX = double.MinValue;
                        for (int i = 0; i < graphData.Dates.Count; i++)
                        {
                            var parsedDate = _parseDate(graphData.Dates[i]);
                            if (parsedDate.HasValue)
                            {
                                var xVal = DateTimeAxis.ToDouble(parsedDate.Value);
                                minX = Math.Min(minX, xVal);
                                maxX = Math.Max(maxX, xVal);
                            }
                        }

                        if (minX != double.MaxValue)
                        {
                            specLine.Points.Add(new DataPoint(minX, graphData.SpecValue.Value));
                            specLine.Points.Add(new DataPoint(maxX, graphData.SpecValue.Value));
                        }

                        model.Series.Add(specLine);
                    }

                    if (graphData.UpperLimitValue.HasValue)
                    {
                        var upperLine = new LineSeries
                        {
                            Title = "Upper Limit",
                            Color = OxyColor.FromArgb(graphData.UpperColor.A, graphData.UpperColor.R,
                                                      graphData.UpperColor.G, graphData.UpperColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        double minX = double.MaxValue;
                        double maxX = double.MinValue;
                        for (int i = 0; i < graphData.Dates.Count; i++)
                        {
                            var parsedDate = _parseDate(graphData.Dates[i]);
                            if (parsedDate.HasValue)
                            {
                                var xVal = DateTimeAxis.ToDouble(parsedDate.Value);
                                minX = Math.Min(minX, xVal);
                                maxX = Math.Max(maxX, xVal);
                            }
                        }

                        if (minX != double.MaxValue)
                        {
                            upperLine.Points.Add(new DataPoint(minX, graphData.UpperLimitValue.Value));
                            upperLine.Points.Add(new DataPoint(maxX, graphData.UpperLimitValue.Value));
                        }

                        model.Series.Add(upperLine);
                    }

                    if (graphData.LowerLimitValue.HasValue)
                    {
                        var lowerLine = new LineSeries
                        {
                            Title = "Lower Limit",
                            Color = OxyColor.FromArgb(graphData.LowerColor.A, graphData.LowerColor.R,
                                                      graphData.LowerColor.G, graphData.LowerColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        double minX = double.MaxValue;
                        double maxX = double.MinValue;
                        for (int i = 0; i < graphData.Dates.Count; i++)
                        {
                            var parsedDate = _parseDate(graphData.Dates[i]);
                            if (parsedDate.HasValue)
                            {
                                var xVal = DateTimeAxis.ToDouble(parsedDate.Value);
                                minX = Math.Min(minX, xVal);
                                maxX = Math.Max(maxX, xVal);
                            }
                        }

                        if (minX != double.MaxValue)
                        {
                            lowerLine.Points.Add(new DataPoint(minX, graphData.LowerLimitValue.Value));
                            lowerLine.Points.Add(new DataPoint(maxX, graphData.LowerLimitValue.Value));
                        }

                        model.Series.Add(lowerLine);
                    }
                }
                else
                {
                    // 순번 기반 SPEC/Limit 선
                    int maxSequence = useFixedSingleX ? 1 : graphData.Dates.Count;

                    if (graphData.SpecValue.HasValue && maxSequence > 0)
                    {
                        var specLine = new LineSeries
                        {
                            Title = "SPEC",
                            Color = OxyColor.FromArgb(graphData.SpecColor.A, graphData.SpecColor.R,
                                                      graphData.SpecColor.G, graphData.SpecColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        specLine.Points.Add(new DataPoint(1, graphData.SpecValue.Value));
                        specLine.Points.Add(new DataPoint(maxSequence, graphData.SpecValue.Value));

                        model.Series.Add(specLine);
                    }

                    if (graphData.UpperLimitValue.HasValue && maxSequence > 0)
                    {
                        var upperLine = new LineSeries
                        {
                            Title = "Upper Limit",
                            Color = OxyColor.FromArgb(graphData.UpperColor.A, graphData.UpperColor.R,
                                                      graphData.UpperColor.G, graphData.UpperColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        upperLine.Points.Add(new DataPoint(1, graphData.UpperLimitValue.Value));
                        upperLine.Points.Add(new DataPoint(maxSequence, graphData.UpperLimitValue.Value));

                        model.Series.Add(upperLine);
                    }

                    if (graphData.LowerLimitValue.HasValue && maxSequence > 0)
                    {
                        var lowerLine = new LineSeries
                        {
                            Title = "Lower Limit",
                            Color = OxyColor.FromArgb(graphData.LowerColor.A, graphData.LowerColor.R,
                                                      graphData.LowerColor.G, graphData.LowerColor.B),
                            StrokeThickness = 2,
                            LineStyle = LineStyle.Dash
                        };

                        lowerLine.Points.Add(new DataPoint(1, graphData.LowerLimitValue.Value));
                        lowerLine.Points.Add(new DataPoint(maxSequence, graphData.LowerLimitValue.Value));

                        model.Series.Add(lowerLine);
                    }
                }
            }

            // 전체 NG Rate를 제목에 추가
            if (totalPoints > 0)
            {
                double totalNgRate = (double)totalNgPoints / totalPoints * 100.0;
                model.Title = $"Daily Sampling Scatter Plot - Total NG: {totalNgPoints}/{totalPoints} (NG Rate: {totalNgRate:F1}%)";
            }

            model.IsLegendVisible = true;
            foreach (var series in model.Series)
            {
                series.RenderInLegend = !string.IsNullOrWhiteSpace(series.Title);
            }
            ApplyScatterVisibility(model);
            return model;
        }

        private void ApplyScatterVisibility(PlotModel model)
        {
            for (int index = model.Series.Count - 1; index >= 0; index--)
            {
                string title = model.Series[index].Title ?? string.Empty;
                if (string.Equals(title, "Sample", StringComparison.OrdinalIgnoreCase) && !ShowSampleSeries)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if ((string.Equals(title, "Avg", StringComparison.OrdinalIgnoreCase) ||
                     title.EndsWith("_avg", StringComparison.OrdinalIgnoreCase)) && !ShowAvgSeries)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if (string.Equals(title, "Max", StringComparison.OrdinalIgnoreCase) && !ShowMaxAnnotations)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if (string.Equals(title, "Min", StringComparison.OrdinalIgnoreCase) && !ShowMinAnnotations)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if (string.Equals(title, "SPEC", StringComparison.OrdinalIgnoreCase) && !ShowSpecSeries)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if (string.Equals(title, "Upper Limit", StringComparison.OrdinalIgnoreCase) && !ShowUpperLimitSeries)
                {
                    model.Series.RemoveAt(index);
                    continue;
                }

                if (string.Equals(title, "Lower Limit", StringComparison.OrdinalIgnoreCase) && !ShowLowerLimitSeries)
                {
                    model.Series.RemoveAt(index);
                }
            }

            for (int index = model.Annotations.Count - 1; index >= 0; index--)
            {
                if (model.Annotations[index] is not TextAnnotation textAnnotation)
                {
                    continue;
                }

                string text = textAnnotation.Text ?? string.Empty;
                if (text.StartsWith("Avg:", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("Max:", StringComparison.OrdinalIgnoreCase) ||
                    text.StartsWith("Min:", StringComparison.OrdinalIgnoreCase))
                {
                    model.Annotations.RemoveAt(index);
                    continue;
                }

                if (text.StartsWith("NG:", StringComparison.OrdinalIgnoreCase) && !ShowAvgSeries)
                {
                    model.Annotations.RemoveAt(index);
                    continue;
                }
            }
        }

        private void AddScatterSeriesByColumn(
            PlotModel model,
            DailySamplingGraphData graphData,
            bool useDate,
            bool useFixedSingleX,
            ref int totalNgPoints,
            ref int totalPoints)
        {
            int rowCount = GetAvailableRowCount(graphData);
            if (rowCount == 0 || graphData.Data == null)
            {
                return;
            }

            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
            {
                string columnName = graphData.Data.Columns[colIndex].ColumnName;
                var scatterSeries = new ScatterSeries
                {
                    Title = columnName,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                        graphData.DataColor.G, graphData.DataColor.B)
                };

                var avgSeries = new LineSeries
                {
                    Title = $"{columnName}_avg",
                    Color = OxyColors.Green,
                    StrokeThickness = 1.5
                };

                if (useDate)
                {
                    var dateStats = new Dictionary<DateTime, List<double>>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        var parsedDate = _parseDate(graphData.Dates[rowIndex]);
                        if (!parsedDate.HasValue)
                        {
                            continue;
                        }

                        var dataRow = graphData.Data.Rows[rowIndex];
                        if (!TryParseNullableDouble(dataRow[colIndex]?.ToString(), out double value))
                        {
                            continue;
                        }

                        var date = parsedDate.Value;
                        var dateValue = DateTimeAxis.ToDouble(date);
                        scatterSeries.Points.Add(new ScatterPoint(dateValue, value));

                        if (!dateStats.TryGetValue(date, out var bucket))
                        {
                            bucket = new List<double>();
                            dateStats[date] = bucket;
                        }
                        bucket.Add(value);

                        var rowLimits = GetLimitsForColumn(graphData, colIndex);
                        if (rowLimits.Upper.HasValue && rowLimits.Lower.HasValue)
                        {
                            totalPoints++;
                            if (value > rowLimits.Upper.Value || value < rowLimits.Lower.Value)
                            {
                                totalNgPoints++;
                            }
                        }
                    }

                    foreach (var kvp in dateStats.OrderBy(x => x.Key))
                    {
                        avgSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(kvp.Key), kvp.Value.Average()));
                    }
                }
                else
                {
                    var sequenceStats = new Dictionary<int, List<double>>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        int sequence = useFixedSingleX ? 1 : rowIndex + 1;
                        var dataRow = graphData.Data.Rows[rowIndex];
                        if (!TryParseNullableDouble(dataRow[colIndex]?.ToString(), out double value))
                        {
                            continue;
                        }

                        scatterSeries.Points.Add(new ScatterPoint(sequence, value));
                        if (!sequenceStats.TryGetValue(sequence, out var bucket))
                        {
                            bucket = new List<double>();
                            sequenceStats[sequence] = bucket;
                        }
                        bucket.Add(value);

                        var rowLimits = GetLimitsForColumn(graphData, colIndex);
                        if (rowLimits.Upper.HasValue && rowLimits.Lower.HasValue)
                        {
                            totalPoints++;
                            if (value > rowLimits.Upper.Value || value < rowLimits.Lower.Value)
                            {
                                totalNgPoints++;
                            }
                        }
                    }

                    foreach (var kvp in sequenceStats.OrderBy(x => x.Key))
                    {
                        avgSeries.Points.Add(new DataPoint(kvp.Key, kvp.Value.Average()));
                    }
                }

                if (scatterSeries.Points.Count > 0)
                {
                    model.Series.Add(scatterSeries);
                }
                if (avgSeries.Points.Count > 0)
                {
                    model.Series.Add(avgSeries);
                }
            }
        }

        private void CreateCpkPlot()
        {
            CpkPlotModel = BuildCpkPlotModel();
        }

        private PlotModel BuildCpkPlotModel()
        {
            var model = new PlotModel { Title = "CPK by Date" };
            ApplyCommonPlotStyle(model);

            // X축 타입 결정 (첫 번째 파일 기준)
            DailySamplingGraphData? firstGraphData = _graphDataList.FirstOrDefault();
            var xAxisMode = firstGraphData?.XAxisMode ?? XAxisMode.Date;
            bool useDate = xAxisMode == XAxisMode.Date;
            bool useFixedSingleX = xAxisMode == XAxisMode.None &&
                                   firstGraphData?.NoXAxisDisplayMode == NoXAxisDisplayMode.FixedSingleX;
            bool splitByColumn = firstGraphData?.DisplayMode == ValuePlotDisplayMode.SplitByColumn;

            // X축 설정
            if (useDate)
            {
                // X축: 날짜 (DateTimeAxis)
                var dateAxis = new DateTimeAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = "Date",
                    StringFormat = "MM-dd",
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IntervalType = DateTimeIntervalType.Days
                };
                model.Axes.Add(dateAxis);
            }
            else
            {
                // X축: 순번 (LinearAxis)
                var sequenceAxis = new LinearAxis
                {
                    Position = AxisPosition.Bottom,
                    Title = useFixedSingleX ? string.Empty :
                        (xAxisMode == XAxisMode.None ? "Sample Order" : "Sequence"),
                    MajorGridlineStyle = LineStyle.Solid,
                    MinorGridlineStyle = LineStyle.Dot,
                    IsAxisVisible = !useFixedSingleX
                };
                model.Axes.Add(sequenceAxis);
            }

            // Y축: CPK 값
            var cpkAxis = new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "CPK",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };
            model.Axes.Add(cpkAxis);

            // 각 파일에 대해 CPK 그래프 생성
            foreach (var graphData in _graphDataList)
            {
                if (splitByColumn)
                {
                    AddCpkSeriesByColumn(model, graphData, useDate, useFixedSingleX);
                    continue;
                }

                // SPEC, Upper, Lower 값이 모두 있어야 CPK 계산 가능
                if (!graphData.SpecValue.HasValue ||
                    !graphData.UpperLimitValue.HasValue ||
                    !graphData.LowerLimitValue.HasValue)
                {
                    continue;
                }

                var cpkSeries = new LineSeries
                {
                    Title = "CPK",
                    Color = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                              graphData.DataColor.G, graphData.DataColor.B),
                    StrokeThickness = 2,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                    graphData.DataColor.G, graphData.DataColor.B)
                };

                if (useDate)
                {
                    // 날짜별로 데이터 그룹화 (중복 날짜 처리)
                    var dateValues = new Dictionary<DateTime, List<double>>();

                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        var dateStr = graphData.Dates[rowIndex];
                        var parsedDate = _parseDate(dateStr);

                        if (parsedDate.HasValue)
                        {
                            if (!dateValues.ContainsKey(parsedDate.Value))
                            {
                                dateValues[parsedDate.Value] = new List<double>();
                            }

                            // 해당 날짜의 모든 샘플 값 수집
                            var dataRow = graphData.Data.Rows[rowIndex];
                            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                            {
                                var valueStr = dataRow[colIndex]?.ToString();
                                if (!string.IsNullOrWhiteSpace(valueStr) &&
                                    double.TryParse(valueStr, out double value))
                                {
                                    dateValues[parsedDate.Value].Add(value);
                                }
                            }
                        }
                    }

                    // 날짜별 CPK 계산
                    foreach (var kvp in dateValues.OrderBy(x => x.Key))
                    {
                        var date = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var dateValue = DateTimeAxis.ToDouble(date);
                            var mean = values.Average();
                            var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
                            var stdDev = Math.Sqrt(variance);

                            if (stdDev > 0)
                            {
                                var usl = graphData.UpperLimitValue.Value;
                                var lsl = graphData.LowerLimitValue.Value;

                                var cpkUpper = (usl - mean) / (3 * stdDev);
                                var cpkLower = (mean - lsl) / (3 * stdDev);
                                var cpk = Math.Min(cpkUpper, cpkLower);

                                cpkSeries.Points.Add(new DataPoint(dateValue, cpk));

                                // CPK 값 텍스트 주석 추가
                                var cpkAnnotation = new TextAnnotation
                                {
                                    Text = $"{cpk:F2}",
                                    TextPosition = new DataPoint(dateValue, cpk),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                                  graphData.DataColor.G, graphData.DataColor.B)
                                };
                                model.Annotations.Add(cpkAnnotation);
                            }
                        }
                    }
                }
                else
                {
                    // 순번별로 데이터 그룹화
                    var sequenceValues = new Dictionary<int, List<double>>();

                    for (int rowIndex = 0; rowIndex < graphData.Dates.Count; rowIndex++)
                    {
                        int sequence = useFixedSingleX ? 1 : rowIndex + 1;

                        if (!sequenceValues.ContainsKey(sequence))
                        {
                            sequenceValues[sequence] = new List<double>();
                        }

                        // 해당 순번의 모든 샘플 값 수집
                        var dataRow = graphData.Data.Rows[rowIndex];
                        for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                        {
                            var valueStr = dataRow[colIndex]?.ToString();
                            if (!string.IsNullOrWhiteSpace(valueStr) &&
                                double.TryParse(valueStr, out double value))
                            {
                                sequenceValues[sequence].Add(value);
                            }
                        }
                    }

                    // 순번별 CPK 계산
                    foreach (var kvp in sequenceValues.OrderBy(x => x.Key))
                    {
                        var sequence = kvp.Key;
                        var values = kvp.Value;

                        if (values.Count > 0)
                        {
                            var mean = values.Average();
                            var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
                            var stdDev = Math.Sqrt(variance);

                            if (stdDev > 0)
                            {
                                var usl = graphData.UpperLimitValue.Value;
                                var lsl = graphData.LowerLimitValue.Value;

                                var cpkUpper = (usl - mean) / (3 * stdDev);
                                var cpkLower = (mean - lsl) / (3 * stdDev);
                                var cpk = Math.Min(cpkUpper, cpkLower);

                                cpkSeries.Points.Add(new DataPoint(sequence, cpk));

                                // CPK 값 텍스트 주석 추가
                                var cpkAnnotation = new TextAnnotation
                                {
                                    Text = $"{cpk:F2}",
                                    TextPosition = new DataPoint(sequence, cpk),
                                    TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                                    TextVerticalAlignment = OxyPlot.VerticalAlignment.Bottom,
                                    Stroke = OxyColors.Transparent,
                                    FontSize = 9,
                                    TextColor = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                                                                  graphData.DataColor.G, graphData.DataColor.B)
                                };
                                model.Annotations.Add(cpkAnnotation);
                            }
                        }
                    }
                }

                model.Series.Add(cpkSeries);
            }

            model.IsLegendVisible = true;
            foreach (var series in model.Series)
            {
                series.RenderInLegend = !string.IsNullOrWhiteSpace(series.Title);
            }
            ApplyCpkVisibility(model);
            return model;
        }

        private void ApplyCpkVisibility(PlotModel model)
        {
            if (ShowCpkSeries)
            {
                return;
            }

            model.Series.Clear();
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
                    LegendOrientation = LegendOrientation.Vertical
                });
            }
        }

        private void AddCpkSeriesByColumn(
            PlotModel model,
            DailySamplingGraphData graphData,
            bool useDate,
            bool useFixedSingleX)
        {
            int rowCount = GetAvailableRowCount(graphData);
            if (rowCount == 0 || graphData.Data == null)
            {
                return;
            }

            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
            {
                var limits = GetLimitsForColumn(graphData, colIndex);
                if (!limits.Spec.HasValue || !limits.Upper.HasValue || !limits.Lower.HasValue)
                {
                    continue;
                }

                string columnName = graphData.Data.Columns[colIndex].ColumnName;
                var cpkSeries = new LineSeries
                {
                    Title = $"{columnName}_cpk",
                    Color = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                        graphData.DataColor.G, graphData.DataColor.B),
                    StrokeThickness = 2,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 4,
                    MarkerFill = OxyColor.FromArgb(graphData.DataColor.A, graphData.DataColor.R,
                        graphData.DataColor.G, graphData.DataColor.B)
                };

                if (useDate)
                {
                    var dateValues = new Dictionary<DateTime, List<double>>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        var parsedDate = _parseDate(graphData.Dates[rowIndex]);
                        if (!parsedDate.HasValue)
                        {
                            continue;
                        }

                        var dataRow = graphData.Data.Rows[rowIndex];
                        if (!TryParseNullableDouble(dataRow[colIndex]?.ToString(), out double value))
                        {
                            continue;
                        }

                        if (!dateValues.TryGetValue(parsedDate.Value, out var bucket))
                        {
                            bucket = new List<double>();
                            dateValues[parsedDate.Value] = bucket;
                        }
                        bucket.Add(value);
                    }

                    foreach (var kvp in dateValues.OrderBy(x => x.Key))
                    {
                        if (TryCalculateCpk(kvp.Value, limits.Upper.Value, limits.Lower.Value, out double cpk))
                        {
                            cpkSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(kvp.Key), cpk));
                        }
                    }
                }
                else
                {
                    var sequenceValues = new Dictionary<int, List<double>>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        int sequence = useFixedSingleX ? 1 : rowIndex + 1;
                        var dataRow = graphData.Data.Rows[rowIndex];
                        if (!TryParseNullableDouble(dataRow[colIndex]?.ToString(), out double value))
                        {
                            continue;
                        }

                        if (!sequenceValues.TryGetValue(sequence, out var bucket))
                        {
                            bucket = new List<double>();
                            sequenceValues[sequence] = bucket;
                        }
                        bucket.Add(value);
                    }

                    foreach (var kvp in sequenceValues.OrderBy(x => x.Key))
                    {
                        if (TryCalculateCpk(kvp.Value, limits.Upper.Value, limits.Lower.Value, out double cpk))
                        {
                            cpkSeries.Points.Add(new DataPoint(kvp.Key, cpk));
                        }
                    }
                }

                if (cpkSeries.Points.Count > 0)
                {
                    model.Series.Add(cpkSeries);
                }
            }
        }

        private static bool TryCalculateCpk(List<double> values, double usl, double lsl, out double cpk)
        {
            cpk = 0;
            if (values == null || values.Count == 0)
            {
                return false;
            }

            double mean = values.Average();
            double variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;
            double stdDev = Math.Sqrt(variance);
            if (stdDev <= 0)
            {
                return false;
            }

            double cpkUpper = (usl - mean) / (3 * stdDev);
            double cpkLower = (mean - lsl) / (3 * stdDev);
            cpk = Math.Min(cpkUpper, cpkLower);
            return true;
        }

        private static (double? Spec, double? Upper, double? Lower) GetLimitsForColumn(
            DailySamplingGraphData graphData,
            int colIndex)
        {
            double? spec = graphData.SpecValue;
            double? upper = graphData.UpperLimitValue;
            double? lower = graphData.LowerLimitValue;

            if (graphData.Data == null || colIndex < 0 || colIndex >= graphData.Data.Columns.Count)
            {
                return (spec, upper, lower);
            }

            string columnName = graphData.Data.Columns[colIndex].ColumnName;
            if (graphData.ColumnLimits != null &&
                graphData.ColumnLimits.TryGetValue(columnName, out var columnLimit))
            {
                if (TryParseNullableDouble(columnLimit.SpecValue, out var specParsed))
                    spec = specParsed;
                if (TryParseNullableDouble(columnLimit.UpperValue, out var upperParsed))
                    upper = upperParsed;
                if (TryParseNullableDouble(columnLimit.LowerValue, out var lowerParsed))
                    lower = lowerParsed;
            }

            return (spec, upper, lower);
        }

        private static bool TryParseNullableDouble(string? text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture, out value))
            {
                return true;
            }

            return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.CurrentCulture, out value);
        }

        private void BuildSummaryRows()
        {
            List<SummaryValueItem> values = CollectSummaryValues();
            SummaryRows = new List<DailySamplingSummaryRow>();

            if (values.Count == 0)
            {
                return;
            }

            SummaryRows.Add(BuildAggregateSummaryRow(GetSummaryPeriodLabel(), values));
        }

        private List<SummaryValueItem> CollectSummaryValues()
        {
            var values = new List<SummaryValueItem>();

            foreach (DailySamplingGraphData graphData in _graphDataList)
            {
                int rowCount = GetAvailableRowCount(graphData);
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    if (!IsRowWithinSelectedPeriod(graphData, rowIndex))
                    {
                        continue;
                    }

                    AppendRowValues(graphData, rowIndex, values);
                }
            }

            return values;
        }

        private bool IsRowWithinSelectedPeriod(DailySamplingGraphData graphData, int rowIndex)
        {
            XAxisMode xAxisMode = _graphDataList.Count > 0 ? _graphDataList[0].XAxisMode : XAxisMode.Date;
            if (xAxisMode != XAxisMode.Date)
            {
                return true;
            }

            if (graphData.Dates == null || rowIndex < 0 || rowIndex >= graphData.Dates.Count)
            {
                return false;
            }

            DateTime? parsedDate = _parseDate(graphData.Dates[rowIndex]);
            if (!parsedDate.HasValue)
            {
                return false;
            }

            DateTime date = parsedDate.Value.Date;
            if (SummaryFromDate.HasValue && date < SummaryFromDate.Value)
            {
                return false;
            }

            if (SummaryToDate.HasValue && date > SummaryToDate.Value)
            {
                return false;
            }

            return true;
        }

        private string GetSummaryPeriodLabel()
        {
            if (SummaryFromDate.HasValue && SummaryToDate.HasValue)
            {
                return $"{SummaryFromDate.Value:yyyy-MM-dd} ~ {SummaryToDate.Value:yyyy-MM-dd}";
            }

            return "All Data";
        }

        private static DailySamplingSummaryRow BuildAggregateSummaryRow(string label, List<SummaryValueItem> values)
        {
            double max = values.Max(v => v.Value);
            double min = values.Min(v => v.Value);
            double avg = values.Average(v => v.Value);
            int totalCount = values.Count;

            int ngCount = values.Count(v =>
                v.Upper.HasValue &&
                v.Lower.HasValue &&
                (v.Value > v.Upper.Value || v.Value < v.Lower.Value));

            double ppm = totalCount > 0
                ? (double)ngCount / totalCount * 1_000_000.0
                : 0.0;

            double? cpk = null;
            List<SummaryValueItem> limitedItems = values.Where(v => v.Upper.HasValue && v.Lower.HasValue).ToList();
            if (limitedItems.Count == totalCount)
            {
                double firstUpper = limitedItems[0].Upper!.Value;
                double firstLower = limitedItems[0].Lower!.Value;
                bool sameLimits = limitedItems.All(v =>
                    Math.Abs(v.Upper!.Value - firstUpper) < 1e-12 &&
                    Math.Abs(v.Lower!.Value - firstLower) < 1e-12);

                if (sameLimits)
                {
                    double variance = values.Sum(v => Math.Pow(v.Value - avg, 2)) / totalCount;
                    double stdDev = Math.Sqrt(variance);
                    if (stdDev > 0)
                    {
                        double cpkUpper = (firstUpper - avg) / (3 * stdDev);
                        double cpkLower = (avg - firstLower) / (3 * stdDev);
                        cpk = Math.Min(cpkUpper, cpkLower);
                    }
                }
            }

            return new DailySamplingSummaryRow
            {
                XAxisSeriesName = label,
                Max = max,
                Min = min,
                Avg = avg,
                Cpk = cpk,
                TotalCount = totalCount,
                NgCount = ngCount,
                Ppm = ppm
            };
        }

        private void BuildDateSummaryRows(
            DailySamplingGraphData graphData,
            int rowCount,
            bool splitByColumn,
            bool hasMultipleFiles,
            List<DailySamplingSummaryRow> summaryRows)
        {
            if (graphData.Data == null)
            {
                return;
            }

            if (splitByColumn)
            {
                for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                {
                    string columnName = graphData.Data.Columns[colIndex].ColumnName;
                    var dateValues = new Dictionary<DateTime, List<SummaryValueItem>>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        var parsedDate = _parseDate(graphData.Dates[rowIndex]);
                        if (!parsedDate.HasValue)
                        {
                            continue;
                        }

                        if (!dateValues.ContainsKey(parsedDate.Value))
                        {
                            dateValues[parsedDate.Value] = new List<SummaryValueItem>();
                        }
                        AppendColumnValue(graphData, rowIndex, colIndex, dateValues[parsedDate.Value]);
                    }

                    foreach (var kvp in dateValues.OrderBy(x => x.Key))
                    {
                        AddSummaryRow(summaryRows, graphData, hasMultipleFiles,
                            $"{kvp.Key:yyyy-MM-dd} | {columnName}", kvp.Value);
                    }
                }
                return;
            }

            var mergedDateValues = new Dictionary<DateTime, List<SummaryValueItem>>();
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var parsedDate = _parseDate(graphData.Dates[rowIndex]);
                if (!parsedDate.HasValue)
                {
                    continue;
                }

                if (!mergedDateValues.ContainsKey(parsedDate.Value))
                {
                    mergedDateValues[parsedDate.Value] = new List<SummaryValueItem>();
                }
                AppendRowValues(graphData, rowIndex, mergedDateValues[parsedDate.Value]);
            }

            foreach (var kvp in mergedDateValues.OrderBy(x => x.Key))
            {
                AddSummaryRow(summaryRows, graphData, hasMultipleFiles,
                    kvp.Key.ToString("yyyy-MM-dd"), kvp.Value);
            }
        }

        private static void BuildFixedXSummaryRows(
            DailySamplingGraphData graphData,
            int rowCount,
            bool splitByColumn,
            bool hasMultipleFiles,
            List<DailySamplingSummaryRow> summaryRows)
        {
            if (graphData.Data == null)
            {
                return;
            }

            if (splitByColumn)
            {
                for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                {
                    string columnName = graphData.Data.Columns[colIndex].ColumnName;
                    var values = new List<SummaryValueItem>();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        AppendColumnValue(graphData, rowIndex, colIndex, values);
                    }

                    AddSummaryRow(summaryRows, graphData, hasMultipleFiles,
                        $"Fixed X | {columnName}", values);
                }
                return;
            }

            var mergedValues = new List<SummaryValueItem>();
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                AppendRowValues(graphData, rowIndex, mergedValues);
            }
            AddSummaryRow(summaryRows, graphData, hasMultipleFiles, "Fixed X", mergedValues);
        }

        private static void BuildSequenceSummaryRows(
            DailySamplingGraphData graphData,
            int rowCount,
            bool splitByColumn,
            bool hasMultipleFiles,
            List<DailySamplingSummaryRow> summaryRows)
        {
            if (graphData.Data == null)
            {
                return;
            }

            if (splitByColumn)
            {
                for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
                {
                    string columnName = graphData.Data.Columns[colIndex].ColumnName;
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        var values = new List<SummaryValueItem>();
                        AppendColumnValue(graphData, rowIndex, colIndex, values);
                        AddSummaryRow(summaryRows, graphData, hasMultipleFiles,
                            $"{rowIndex + 1} | {columnName}", values);
                    }
                }
                return;
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                var values = new List<SummaryValueItem>();
                AppendRowValues(graphData, rowIndex, values);
                AddSummaryRow(summaryRows, graphData, hasMultipleFiles,
                    (rowIndex + 1).ToString(), values);
            }
        }

        private static int GetAvailableRowCount(DailySamplingGraphData graphData)
        {
            if (graphData.Data == null || graphData.Dates == null)
            {
                return 0;
            }

            return Math.Min(graphData.Data.Rows.Count, graphData.Dates.Count);
        }

        private static void AppendRowValues(
            DailySamplingGraphData graphData,
            int rowIndex,
            List<SummaryValueItem> values)
        {
            if (graphData.Data == null || rowIndex < 0 || rowIndex >= graphData.Data.Rows.Count)
            {
                return;
            }

            for (int colIndex = 1; colIndex < graphData.Data.Columns.Count; colIndex++)
            {
                AppendColumnValue(graphData, rowIndex, colIndex, values);
            }
        }

        private static void AddSummaryRow(
            List<DailySamplingSummaryRow> summaryRows,
            DailySamplingGraphData graphData,
            bool hasMultipleFiles,
            string xAxisLabel,
            List<SummaryValueItem> values)
        {
            if (values.Count == 0)
            {
                return;
            }

            double max = values.Max(v => v.Value);
            double min = values.Min(v => v.Value);
            double avg = values.Average(v => v.Value);
            int totalCount = values.Count;

            int ngCount = values.Count(v =>
                v.Upper.HasValue &&
                v.Lower.HasValue &&
                (v.Value > v.Upper.Value || v.Value < v.Lower.Value));

            double ppm = totalCount > 0
                ? (double)ngCount / totalCount * 1_000_000.0
                : 0.0;

            double? cpk = null;
            var limitedItems = values.Where(v => v.Upper.HasValue && v.Lower.HasValue).ToList();
            if (limitedItems.Count == totalCount)
            {
                double firstUpper = limitedItems[0].Upper!.Value;
                double firstLower = limitedItems[0].Lower!.Value;
                bool sameLimits = limitedItems.All(v =>
                    Math.Abs(v.Upper!.Value - firstUpper) < 1e-12 &&
                    Math.Abs(v.Lower!.Value - firstLower) < 1e-12);

                if (sameLimits)
                {
                    double variance = values.Sum(v => Math.Pow(v.Value - avg, 2)) / totalCount;
                    double stdDev = Math.Sqrt(variance);
                    if (stdDev > 0)
                    {
                        double cpkUpper = (firstUpper - avg) / (3 * stdDev);
                        double cpkLower = (avg - firstLower) / (3 * stdDev);
                        cpk = Math.Min(cpkUpper, cpkLower);
                    }
                }
            }

            string axisSeriesName = hasMultipleFiles
                ? $"{graphData.FileName} | {xAxisLabel}"
                : xAxisLabel;

            summaryRows.Add(new DailySamplingSummaryRow
            {
                XAxisSeriesName = axisSeriesName,
                Max = max,
                Min = min,
                Avg = avg,
                Cpk = cpk,
                TotalCount = totalCount,
                NgCount = ngCount,
                Ppm = ppm
            });
        }

        private static void AppendColumnValue(
            DailySamplingGraphData graphData,
            int rowIndex,
            int colIndex,
            List<SummaryValueItem> values)
        {
            if (graphData.Data == null ||
                rowIndex < 0 || rowIndex >= graphData.Data.Rows.Count ||
                colIndex < 1 || colIndex >= graphData.Data.Columns.Count)
            {
                return;
            }

            var dataRow = graphData.Data.Rows[rowIndex];
            if (!TryParseNullableDouble(dataRow[colIndex]?.ToString(), out double value))
            {
                return;
            }

            var limits = GetLimitsForColumn(graphData, colIndex);
            values.Add(new SummaryValueItem
            {
                Value = value,
                Upper = limits.Upper,
                Lower = limits.Lower
            });
        }
        private void CopySummaryButton_Click(object sender, RoutedEventArgs e)
        {
            if (SummaryDataGrid.Items.Count == 0)
            {
                MessageBox.Show("No statistics data to copy.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("X-axis Series\tMax\tMin\tAvg\tCPK\tTotal Count\tNG Count\tPPM");

            foreach (var item in SummaryDataGrid.Items)
            {
                if (item is not DailySamplingSummaryRow row)
                {
                    continue;
                }

                string cpkText = row.Cpk.HasValue
                    ? row.Cpk.Value.ToString("F2", CultureInfo.InvariantCulture)
                    : "-";

                sb.Append(row.XAxisSeriesName)
                  .Append('\t')
                  .Append(row.Max.ToString("F2", CultureInfo.InvariantCulture))
                  .Append('\t')
                  .Append(row.Min.ToString("F2", CultureInfo.InvariantCulture))
                  .Append('\t')
                  .Append(row.Avg.ToString("F2", CultureInfo.InvariantCulture))
                  .Append('\t')
                  .Append(cpkText)
                  .Append('\t')
                  .Append(row.TotalCount.ToString(CultureInfo.InvariantCulture))
                  .Append('\t')
                  .Append(row.NgCount.ToString(CultureInfo.InvariantCulture))
                  .Append('\t')
                  .Append(row.Ppm.ToString("F0", CultureInfo.InvariantCulture))
                  .AppendLine();
            }

            try
            {
                Clipboard.SetText(sb.ToString(), TextDataFormat.UnicodeText);
                MessageBox.Show("Copied the entire table to the clipboard.\nPaste it in Excel with Ctrl+V.", "Done",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error while copying to clipboard.\n{ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LegendOptionChanged(object sender, RoutedEventArgs e)
        {
            if (_isInitializingView || !IsLoaded || _graphDataList.Count == 0)
            {
                return;
            }

            RefreshView();
        }

        private void SummaryPeriodChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isInitializingView)
            {
                return;
            }

            if (SummaryFromDate.HasValue && SummaryToDate.HasValue && SummaryFromDate > SummaryToDate)
            {
                _isInitializingView = true;
                try
                {
                    if (ReferenceEquals(sender, SummaryFromDatePicker))
                    {
                        SummaryToDatePicker.SelectedDate = SummaryFromDatePicker.SelectedDate;
                    }
                    else
                    {
                        SummaryFromDatePicker.SelectedDate = SummaryToDatePicker.SelectedDate;
                    }
                }
                finally
                {
                    _isInitializingView = false;
                }
            }

            BuildSummaryRows();
            SummaryDataGrid.ItemsSource = null;
            SummaryDataGrid.ItemsSource = SummaryRows;
        }
    }
}
