using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;

namespace GraphMaker
{
    internal static class GraphSampleModelFactory
    {
        public static PlotModel CreateScatterLineSample()
        {
            var model = CreateBaseModel("Sample Line Plot", "Sample #", "Value");
            var line = new LineSeries { Color = OxyColor.FromRgb(0x2B, 0x6C, 0xB0), StrokeThickness = 2 };

            for (int i = 0; i < 12; i++)
            {
                double y = 48 + (i * 2.8) + (i % 2 == 0 ? 2.5 : -1.7);
                line.Points.Add(new DataPoint(i + 1, y));
            }

            model.Series.Add(line);
            return model;
        }

        public static PlotModel CreateScatterPointSample()
        {
            var model = CreateBaseModel("Sample Scatter Plot", "X", "Y");
            var scatter = new ScatterSeries
            {
                MarkerType = MarkerType.Circle,
                MarkerFill = OxyColor.FromRgb(0x1F, 0x78, 0xD1),
                MarkerSize = 3.5
            };

            for (int i = 0; i < 40; i++)
            {
                double x = i * 0.25;
                double y = (x * 1.6) + ((i % 7) - 3) * 0.6 + 8;
                scatter.Points.Add(new ScatterPoint(x, y));
            }

            model.Series.Add(scatter);
            return model;
        }

        public static PlotModel CreateAverageSample()
        {
            var model = CreateBaseModel("Sample X Avg", "Lot", "Avg");
            var series = new LineSeries { Color = OxyColor.FromRgb(0x2E, 0x8B, 0x57), StrokeThickness = 2.2 };
            double[] values = { 11.4, 10.8, 12.1, 11.6, 12.5, 12.0 };

            for (int i = 0; i < values.Length; i++)
            {
                series.Points.Add(new DataPoint(i + 1, values[i]));
            }

            model.Series.Add(series);
            return model;
        }

        public static PlotModel CreateCpkSample()
        {
            var model = CreateBaseModel("Sample X CPK", "Lot", "CPK");
            var series = new LineSeries
            {
                Color = OxyColor.FromRgb(0x4B, 0x86, 0xC8),
                StrokeThickness = 2,
                MarkerType = MarkerType.Square,
                MarkerFill = OxyColor.FromRgb(0x2E, 0x67, 0xA7),
                MarkerSize = 3
            };

            double[] values = { 1.22, 1.08, 1.31, 0.97, 1.18, 1.26 };
            for (int i = 0; i < values.Length; i++)
            {
                series.Points.Add(new DataPoint(i + 1, values[i]));
            }

            model.Series.Add(series);
            return model;
        }

        public static PlotModel CreateValuePlotSample()
        {
            var model = CreateBaseModel("Sample Value Plot", "Date Index", "Measurement");
            var series = new LineSeries
            {
                Color = OxyColor.FromRgb(0x2E, 0x67, 0xA7),
                StrokeThickness = 2,
                MarkerType = MarkerType.Circle,
                MarkerSize = 2.8
            };

            double[] data = { 102.2, 101.8, 103.4, 102.9, 101.3, 100.8, 102.5, 103.0, 102.1, 101.7 };
            for (int i = 0; i < data.Length; i++)
            {
                series.Points.Add(new DataPoint(i + 1, data[i]));
            }

            model.Series.Add(series);
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Horizontal,
                Y = 103.5,
                Color = OxyColor.FromRgb(0xD1, 0x34, 0x38),
                Text = "Upper"
            });
            model.Annotations.Add(new LineAnnotation
            {
                Type = LineAnnotationType.Horizontal,
                Y = 100.5,
                Color = OxyColor.FromRgb(0xD1, 0x34, 0x38),
                Text = "Lower"
            });

            return model;
        }

        public static PlotModel CreateValueMultiSample()
        {
            var model = CreateBaseModel("Sample Multi-Column Plot", "Sample #", "Value");
            var a = new LineSeries { Title = "Column A", Color = OxyColor.FromRgb(0x2B, 0x6C, 0xB0) };
            var b = new LineSeries { Title = "Column B", Color = OxyColor.FromRgb(0x2E, 0x8B, 0x57) };
            var c = new LineSeries { Title = "Column C", Color = OxyColor.FromRgb(0xC2, 0x41, 0x0C) };

            for (int i = 0; i < 10; i++)
            {
                a.Points.Add(new DataPoint(i + 1, 50 + i * 1.2 + (i % 2 == 0 ? 1.1 : -0.8)));
                b.Points.Add(new DataPoint(i + 1, 46 + i * 1.4 + (i % 3 == 0 ? 0.9 : -0.6)));
                c.Points.Add(new DataPoint(i + 1, 52 + i * 1.0 + (i % 4 == 0 ? 1.4 : -0.5)));
            }

            model.Series.Add(a);
            model.Series.Add(b);
            model.Series.Add(c);
            return model;
        }

        public static PlotModel CreateProcessTrendSample()
        {
            var model = CreateBaseModel("Sample Process Trend", "Process A", "Process B");
            var scatter = new ScatterSeries
            {
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = OxyColor.FromRgb(0x1F, 0x78, 0xD1)
            };

            for (int i = 0; i < 36; i++)
            {
                double x = 80 + i * 0.55;
                double y = 18 + x * 0.22 + ((i % 5) - 2) * 0.4;
                scatter.Points.Add(new ScatterPoint(x, y));
            }

            var regression = new LineSeries
            {
                Title = "Trend",
                Color = OxyColor.FromRgb(0xD1, 0x34, 0x38),
                StrokeThickness = 2
            };
            regression.Points.Add(new DataPoint(80, 35.6));
            regression.Points.Add(new DataPoint(99.5, 39.9));

            model.Series.Add(scatter);
            model.Series.Add(regression);
            return model;
        }

        public static PlotModel CreateHeatMapSample()
        {
            var model = new PlotModel { Title = "Sample Heatmap" };
            model.Axes.Add(new LinearColorAxis
            {
                Position = AxisPosition.Right,
                Palette = OxyPalettes.BlueWhiteRed(160),
                Title = "Score"
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Bottom,
                Minimum = 0,
                Maximum = 6,
                MajorStep = 1,
                Title = "Condition 1"
            });
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Minimum = 0,
                Maximum = 5,
                MajorStep = 1,
                Title = "Condition 2"
            });

            var data = new double[5, 6]
            {
                { 12, 18, 21, 17, 13, 10 },
                { 15, 24, 31, 29, 19, 12 },
                { 11, 22, 35, 38, 27, 16 },
                { 9, 17, 26, 30, 25, 18 },
                { 6, 12, 18, 21, 20, 14 }
            };

            model.Series.Add(new HeatMapSeries
            {
                X0 = 0,
                X1 = 6,
                Y0 = 0,
                Y1 = 5,
                Data = data,
                Interpolate = false,
                RenderMethod = HeatMapRenderMethod.Rectangles
            });

            return model;
        }

        private static PlotModel CreateBaseModel(string title, string xTitle, string yTitle)
        {
            var model = new PlotModel { Title = title };
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = xTitle });
            model.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = yTitle });
            return model;
        }
    }
}
