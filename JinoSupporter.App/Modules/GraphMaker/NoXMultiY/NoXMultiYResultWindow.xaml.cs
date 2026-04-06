using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Legends;
using OxyPlot.Series;

namespace GraphMaker
{
    public partial class NoXMultiYResultWindow : Window
    {
        private readonly NoXMultiYGraphResult _result;
        private readonly string _fileName;

        public NoXMultiYResultWindow(string fileName, NoXMultiYGraphResult result)
        {
            InitializeComponent();
            _fileName = fileName;
            _result = result;
            HeaderTextBlock.Text = $"{fileName} - SingleX(No) / Multi Y";
            PlotViewControl.Model = BuildPlotModel();
            StatsGrid.ItemsSource = _result.Columns;
        }

        private void RecommendLimitsButton_Click(object sender, RoutedEventArgs e)
        {
            var window = new NoXMultiYLimitRecommendationWindow(_fileName, _result)
            {
                Owner = this
            };
            window.ShowDialog();
        }

        private PlotModel BuildPlotModel()
        {
            var model = new PlotModel
            {
                Title = "Header-based Scatter Comparison",
                Background = OxyColors.White,
                PlotAreaBackground = OxyColors.White,
                IsLegendVisible = true
            };

            model.TextColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.TitleColor = OxyColor.FromRgb(0x1D, 0x27, 0x33);
            model.PlotAreaBorderColor = OxyColor.FromRgb(0xB8, 0xC7, 0xDB);
            model.Legends.Add(new Legend
            {
                LegendPlacement = LegendPlacement.Outside,
                LegendPosition = LegendPosition.RightTop,
                LegendOrientation = LegendOrientation.Vertical
            });

            var xAxis = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Header",
                GapWidth = 0.2,
                Angle = 25,
                MajorGridlineStyle = LineStyle.None,
                MinorGridlineStyle = LineStyle.None
            };

            foreach (string category in _result.Categories)
            {
                xAxis.Labels.Add(category);
            }

            model.Axes.Add(xAxis);
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Y",
                AxislineStyle = LineStyle.Solid,
                AxislineThickness = 1,
                TickStyle = TickStyle.Outside,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.None
            });

            foreach (var column in _result.Columns)
            {
                var series = new ScatterSeries
                {
                    Title = column.ColumnName,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3.5
                };

                foreach (NoXMultiYPoint point in column.Points)
                {
                    series.Points.Add(new ScatterPoint(ApplyDeterministicJitter(point.X, point.RowIndex), point.Y));
                }

                model.Series.Add(series);
                AddNormalDistributionOverlay(model, column);
            }

            AddPerColumnLimitSeries(model, _result.Columns, column => column.Spec, "Spec", OxyColors.DarkOrange);
            AddPerColumnLimitSeries(model, _result.Columns, column => column.Upper, "USL", OxyColors.IndianRed);
            AddPerColumnLimitSeries(model, _result.Columns, column => column.Lower, "LSL", OxyColors.SeaGreen);

            return model;
        }

        private static double ApplyDeterministicJitter(double baseX, int rowIndex)
        {
            int jitterSeed = ((rowIndex % 9) - 4);
            return baseX + (jitterSeed * 0.035);
        }

        private static void AddPerColumnLimitSeries(
            PlotModel model,
            List<NoXMultiYColumnResult> columns,
            Func<NoXMultiYColumnResult, double?> selector,
            string title,
            OxyColor color)
        {
            var limitSeries = new LineSeries
            {
                Title = title,
                Color = color,
                StrokeThickness = 1.5,
                LineStyle = LineStyle.Dash,
                MarkerType = MarkerType.None
            };

            bool hasPoint = false;
            for (int index = 0; index < columns.Count; index++)
            {
                double? value = selector(columns[index]);
                if (!value.HasValue)
                {
                    limitSeries.Points.Add(DataPoint.Undefined);
                    continue;
                }

                hasPoint = true;
                limitSeries.Points.Add(new DataPoint(index, value.Value));
            }

            if (hasPoint)
            {
                model.Series.Add(limitSeries);
            }
        }

        private static void AddNormalDistributionOverlay(PlotModel model, NoXMultiYColumnResult column)
        {
            if (column.Values.Count < 2 || column.StdDev <= 0)
            {
                return;
            }

            int categoryIndex = (int)Math.Round(column.Points.FirstOrDefault()?.X ?? 0);
            double minY = column.Min;
            double maxY = column.Max;
            if (Math.Abs(maxY - minY) < 0.0000001)
            {
                return;
            }

            const int sampleCount = 48;
            const double maxHalfWidth = 0.28;
            double peakPdf = 1.0 / (column.StdDev * Math.Sqrt(2.0 * Math.PI));

            var distributionSeries = new LineSeries
            {
                Title = $"{column.ColumnName} Normal",
                Color = OxyColor.FromAColor(120, OxyColors.DimGray),
                StrokeThickness = 1.2,
                LineStyle = LineStyle.Solid,
                RenderInLegend = false
            };

            var centerSeries = new LineSeries
            {
                Color = OxyColor.FromAColor(140, OxyColors.DimGray),
                StrokeThickness = 1.2,
                LineStyle = LineStyle.Dash,
                RenderInLegend = false
            };

            var axisSeries = new LineSeries
            {
                Color = OxyColor.FromAColor(90, OxyColors.Gray),
                StrokeThickness = 1.0,
                LineStyle = LineStyle.Solid,
                RenderInLegend = false
            };

            for (int i = 0; i < sampleCount; i++)
            {
                double y = minY + ((maxY - minY) * i / (sampleCount - 1));
                double z = (y - column.Avg) / column.StdDev;
                double pdf = peakPdf * Math.Exp(-0.5 * z * z);
                double width = peakPdf > 0 ? (pdf / peakPdf) * maxHalfWidth : 0;

                distributionSeries.Points.Add(new DataPoint(categoryIndex + width, y));
            }

            axisSeries.Points.Add(new DataPoint(categoryIndex, minY));
            axisSeries.Points.Add(new DataPoint(categoryIndex, maxY));
            centerSeries.Points.Add(new DataPoint(categoryIndex - 0.22, column.Avg));
            centerSeries.Points.Add(new DataPoint(categoryIndex + 0.22, column.Avg));

            model.Series.Add(axisSeries);
            model.Series.Add(distributionSeries);
            model.Series.Add(centerSeries);
            model.Annotations.Add(new TextAnnotation
            {
                Text = column.Avg.ToString("F3"),
                TextPosition = new DataPoint(categoryIndex + 0.24, column.Avg),
                Stroke = OxyColors.Transparent,
                TextColor = OxyColor.FromAColor(180, OxyColors.DimGray),
                FontSize = 10,
                TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Left,
                TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle
            });
        }
    }
}
