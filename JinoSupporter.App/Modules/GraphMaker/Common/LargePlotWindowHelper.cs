using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Wpf;
using Color = System.Windows.Media.Color;

namespace GraphMaker
{
    internal static class LargePlotWindowHelper
    {
        private sealed class ColorChoice
        {
            public string Name { get; init; } = string.Empty;
            public Color Color { get; init; }
        }

        private sealed class SeriesColorRow
        {
            public Series Series { get; init; } = null!;
            public ComboBox ComboBox { get; init; } = null!;
        }

        private sealed class SeriesVisibilityRow
        {
            public Series Series { get; init; } = null!;
            public CheckBox CheckBox { get; init; } = null!;
        }

        private static readonly IReadOnlyList<ColorChoice> ColorChoices = new[]
        {
            new ColorChoice { Name = "Black", Color = Colors.Black },
            new ColorChoice { Name = "Blue", Color = Colors.DodgerBlue },
            new ColorChoice { Name = "Red", Color = Colors.IndianRed },
            new ColorChoice { Name = "Green", Color = Colors.SeaGreen },
            new ColorChoice { Name = "Orange", Color = Colors.DarkOrange },
            new ColorChoice { Name = "Purple", Color = Colors.MediumPurple },
            new ColorChoice { Name = "Gray", Color = Colors.Gray }
        };

        public static Window CreateLargePlotWindow(Window owner, string title, PlotModel model)
        {
            var primaryPlotView = new PlotView
            {
                Model = model
            };
            var normalDistributionView = new PlotView
            {
                Model = BuildNormalDistributionModel(title, model)
            };
            var lowerLeftPlotView = new PlotView
            {
                Model = ClonePlotModel(model)
            };
            var lowerRightPlotView = new PlotView
            {
                Model = ClonePlotModel(model)
            };
            var plotViews = new[] { primaryPlotView, normalDistributionView, lowerLeftPlotView, lowerRightPlotView };

            var visibilityPanel = new StackPanel();

            var legendColorButton = new Button
            {
                Content = "Legend Colors",
                Width = 120,
                Height = 30,
                Margin = new Thickness(0, 0, 8, 8),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };

            var closeButton = new Button
            {
                Content = "Close",
                Width = 88,
                Height = 30,
                Margin = new Thickness(0, 0, 0, 8),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            };

            var buttonGrid = new Grid();
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            buttonGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(legendColorButton, 0);
            Grid.SetColumn(closeButton, 2);
            buttonGrid.Children.Add(legendColorButton);
            buttonGrid.Children.Add(closeButton);

            var root = new Grid();
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            root.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(280) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var plotHost = new Grid();
            plotHost.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            plotHost.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            Grid.SetRow(buttonGrid, 0);
            var plotGrid = CreateQuarterPlotGrid(title, primaryPlotView, normalDistributionView, lowerLeftPlotView, lowerRightPlotView);
            Grid.SetRow(plotGrid, 1);
            plotHost.Children.Add(buttonGrid);
            plotHost.Children.Add(plotGrid);

            var sidePanel = new Border
            {
                Margin = new Thickness(12, 0, 0, 0),
                BorderBrush = new SolidColorBrush(Color.FromRgb(208, 215, 222)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(10)
            };

            var sideContent = new Grid();
            sideContent.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            sideContent.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            var sideHeader = new TextBlock
            {
                Text = "Visible Elements",
                FontWeight = System.Windows.FontWeights.SemiBold
            };

            var sideScroll = new ScrollViewer
            {
                Margin = new Thickness(0, 10, 0, 0),
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = visibilityPanel
            };

            Grid.SetRow(sideHeader, 0);
            Grid.SetRow(sideScroll, 1);
            sideContent.Children.Add(sideHeader);
            sideContent.Children.Add(sideScroll);
            sidePanel.Child = sideContent;

            Grid.SetColumn(plotHost, 0);
            Grid.SetColumn(sidePanel, 1);
            Grid.SetRowSpan(plotHost, 2);
            Grid.SetRowSpan(sidePanel, 2);
            root.Children.Add(plotHost);
            root.Children.Add(sidePanel);

            var window = new Window
            {
                Title = $"{title} - Large View",
                Content = root,
                Width = 1500,
                Height = 900,
                MinWidth = 900,
                MinHeight = 600,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner
            };

            legendColorButton.Click += (_, _) => OpenLegendColorEditor(window, primaryPlotView);
            closeButton.Click += (_, _) => window.Close();
            PopulateSeriesVisibilityPanel(plotViews, visibilityPanel);
            window.Closed += (_, _) =>
            {
                foreach (var plotView in plotViews)
                {
                    plotView.Model = null;
                }
            };
            return window;
        }

        public static void PopulateSeriesVisibilityPanel(PlotView plotView, Panel panel)
        {
            PopulateSeriesVisibilityPanel(new[] { plotView }, panel);
        }

        private static void PopulateSeriesVisibilityPanel(IReadOnlyList<PlotView> plotViews, Panel panel)
        {
            panel.Children.Clear();

            var primaryPlotView = plotViews.FirstOrDefault();
            if (primaryPlotView?.Model == null)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "No plot model loaded.",
                    TextWrapping = TextWrapping.Wrap
                });
                return;
            }

            var rows = new List<SeriesVisibilityRow>();
            for (int seriesIndex = 0; seriesIndex < primaryPlotView.Model.Series.Count; seriesIndex++)
            {
                var series = primaryPlotView.Model.Series[seriesIndex];
                string title = !string.IsNullOrWhiteSpace(series.Title)
                    ? series.Title
                    : $"Series {rows.Count + 1}";

                var checkBox = new CheckBox
                {
                    Content = title,
                    IsChecked = series.IsVisible,
                    Margin = new Thickness(0, 0, 0, 6)
                };
                checkBox.Checked += (_, _) =>
                {
                    SetSeriesVisibility(plotViews, seriesIndex, true);
                };
                checkBox.Unchecked += (_, _) =>
                {
                    SetSeriesVisibility(plotViews, seriesIndex, false);
                };

                rows.Add(new SeriesVisibilityRow { Series = series, CheckBox = checkBox });
                panel.Children.Add(checkBox);
            }

            if (rows.Count == 0)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = "No graph elements available.",
                    TextWrapping = TextWrapping.Wrap
                });
            }
        }

        private static void SetSeriesVisibility(IReadOnlyList<PlotView> plotViews, int seriesIndex, bool isVisible)
        {
            foreach (var plotView in plotViews)
            {
                if (plotView.Model == null || seriesIndex >= plotView.Model.Series.Count)
                {
                    continue;
                }

                plotView.Model.Series[seriesIndex].IsVisible = isVisible;
                plotView.InvalidatePlot(false);
            }
        }

        public static void OpenLegendColorEditor(Window owner, PlotView plotView)
        {
            if (plotView.Model == null)
            {
                return;
            }

            var legendSeries = plotView.Model.Series
                .Where(s => s.RenderInLegend && !string.IsNullOrWhiteSpace(s.Title))
                .ToList();
            if (legendSeries.Count == 0)
            {
                MessageBox.Show(owner, "No legend series are available for color editing.",
                    "Legend Colors", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var rows = new List<SeriesColorRow>();
            var stack = new StackPanel();

            foreach (var series in legendSeries)
            {
                var rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });

                var titleBlock = new TextBlock
                {
                    Text = series.Title,
                    VerticalAlignment = System.Windows.VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0),
                    TextWrapping = TextWrapping.Wrap
                };

                var comboBox = new ComboBox
                {
                    ItemsSource = ColorChoices,
                    DisplayMemberPath = nameof(ColorChoice.Name),
                    SelectedIndex = Math.Max(0, GetColorChoiceIndex(GetSeriesColor(series)))
                };

                Grid.SetColumn(titleBlock, 0);
                Grid.SetColumn(comboBox, 1);
                rowGrid.Children.Add(titleBlock);
                rowGrid.Children.Add(comboBox);
                stack.Children.Add(rowGrid);
                rows.Add(new SeriesColorRow { Series = series, ComboBox = comboBox });
            }

            var saveButton = new Button
            {
                Content = "Apply",
                Width = 90,
                Height = 30,
                Margin = new Thickness(0, 0, 0, 0),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            };
            var cancelButton = new Button
            {
                Content = "Cancel",
                Width = 90,
                Height = 30,
                Margin = new Thickness(0, 0, 8, 0),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            };

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            buttonPanel.Children.Add(cancelButton);
            buttonPanel.Children.Add(saveButton);

            var content = new StackPanel { Margin = new Thickness(12) };
            content.Children.Add(new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                Content = stack,
                MaxHeight = 480
            });
            content.Children.Add(buttonPanel);

            var dialog = new Window
            {
                Title = "Legend Colors",
                Content = content,
                Width = 480,
                Height = 620,
                MinWidth = 420,
                MinHeight = 300,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = owner,
                ResizeMode = ResizeMode.CanResize
            };

            cancelButton.Click += (_, _) => dialog.Close();
            saveButton.Click += (_, _) =>
            {
                foreach (var row in rows)
                {
                    if (row.ComboBox.SelectedItem is not ColorChoice choice)
                    {
                        continue;
                    }

                    ApplyColorToSeries(row.Series, choice.Color);
                }

                plotView.InvalidatePlot(true);
                dialog.Close();
            };

            dialog.ShowDialog();
        }

        private static int GetColorChoiceIndex(OxyColor color)
        {
            for (int i = 0; i < ColorChoices.Count; i++)
            {
                var choice = ColorChoices[i].Color;
                if (choice.A == color.A && choice.R == color.R && choice.G == color.G && choice.B == color.B)
                {
                    return i;
                }
            }

            return 0;
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

        private static void ApplyColorToSeries(Series series, Color color)
        {
            var oxyColor = OxyColor.FromArgb(color.A, color.R, color.G, color.B);
            switch (series)
            {
                case LineSeries line:
                    line.Color = oxyColor;
                    if (line.MarkerType != MarkerType.None)
                    {
                        line.MarkerFill = oxyColor;
                    }
                    break;

                case ScatterSeries scatter:
                    scatter.MarkerFill = oxyColor;
                    scatter.MarkerStroke = oxyColor;
                    break;
            }
        }

        private static Grid CreateQuarterPlotGrid(
            string title,
            PlotView primaryPlotView,
            PlotView normalDistributionView,
            PlotView lowerLeftPlotView,
            PlotView lowerRightPlotView)
        {
            var grid = new Grid();
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            AddPlotTile(grid, 0, 0, title, primaryPlotView, new Thickness(0, 0, 8, 8));
            AddPlotTile(grid, 0, 1, "Normal Distribution", normalDistributionView, new Thickness(8, 0, 0, 8));
            AddPlotTile(grid, 1, 0, $"{title} Copy 1", lowerLeftPlotView, new Thickness(0, 8, 8, 0));
            AddPlotTile(grid, 1, 1, $"{title} Copy 2", lowerRightPlotView, new Thickness(8, 8, 0, 0));
            return grid;
        }

        private static void AddPlotTile(Grid host, int row, int column, string header, PlotView plotView, Thickness margin)
        {
            var groupBox = new GroupBox
            {
                Header = header,
                Margin = margin,
                Content = plotView
            };
            Grid.SetRow(groupBox, row);
            Grid.SetColumn(groupBox, column);
            host.Children.Add(groupBox);
        }

        private static PlotModel BuildNormalDistributionModel(string title, PlotModel source)
        {
            var model = new PlotModel
            {
                Title = $"{title} Normal Distribution",
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

            foreach (var series in source.Series)
            {
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
                var distribution = new LineSeries
                {
                    Title = string.IsNullOrWhiteSpace(series.Title) ? "Distribution" : series.Title,
                    RenderInLegend = true,
                    StrokeThickness = 2,
                    Color = GetSeriesColor(series)
                };

                const int steps = 100;
                for (int i = 0; i <= steps; i++)
                {
                    double x = min + ((max - min) * i / steps);
                    double y = (1.0 / (stdDev * Math.Sqrt(2 * Math.PI))) *
                               Math.Exp(-0.5 * Math.Pow((x - mean) / stdDev, 2));
                    distribution.Points.Add(new DataPoint(x, y));
                }

                model.Series.Add(distribution);
            }

            return model;
        }

        private static List<double> ExtractSeriesValues(Series series)
        {
            return series switch
            {
                LineSeries line => line.Points.Select(point => point.Y).Where(value => !double.IsNaN(value) && !double.IsInfinity(value)).ToList(),
                ScatterSeries scatter => scatter.Points.Select(point => point.Y).Where(value => !double.IsNaN(value) && !double.IsInfinity(value)).ToList(),
                _ => new List<double>()
            };
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
                if (axis is LinearAxis linearAxis)
                {
                    clone.Axes.Add(new LinearAxis
                    {
                        Position = linearAxis.Position,
                        Title = linearAxis.Title,
                        Minimum = linearAxis.Minimum,
                        Maximum = linearAxis.Maximum,
                        MajorGridlineStyle = linearAxis.MajorGridlineStyle,
                        MinorGridlineStyle = linearAxis.MinorGridlineStyle
                    });
                }
            }

            foreach (var series in source.Series)
            {
                switch (series)
                {
                    case LineSeries line:
                        var clonedLine = new LineSeries
                        {
                            Title = line.Title,
                            Color = line.Color,
                            StrokeThickness = line.StrokeThickness,
                            LineStyle = line.LineStyle,
                            MarkerType = line.MarkerType,
                            MarkerSize = line.MarkerSize,
                            MarkerFill = line.MarkerFill,
                            RenderInLegend = line.RenderInLegend,
                            IsVisible = line.IsVisible
                        };
                        foreach (var point in line.Points)
                        {
                            clonedLine.Points.Add(point);
                        }
                        clone.Series.Add(clonedLine);
                        break;

                    case ScatterSeries scatter:
                        var clonedScatter = new ScatterSeries
                        {
                            Title = scatter.Title,
                            MarkerType = scatter.MarkerType,
                            MarkerSize = scatter.MarkerSize,
                            MarkerFill = scatter.MarkerFill,
                            MarkerStroke = scatter.MarkerStroke,
                            MarkerStrokeThickness = scatter.MarkerStrokeThickness,
                            RenderInLegend = scatter.RenderInLegend,
                            IsVisible = scatter.IsVisible
                        };
                        foreach (var point in scatter.Points)
                        {
                            clonedScatter.Points.Add(new ScatterPoint(point.X, point.Y, point.Size, point.Value, point.Tag));
                        }
                        clone.Series.Add(clonedScatter);
                        break;
                }
            }

            return clone;
        }
    }
}
