using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using Color = System.Windows.Media.Color;
using Colors = System.Windows.Media.Colors;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GraphMaker
{
    public partial class DailyDataTrendView : UserControl, INotifyPropertyChanged
    {
        private sealed class DailyDataTrendReportState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public int SavedPlotColorIndex { get; set; }
            public RawTableData RawData { get; set; } = new();
            public Dictionary<string, string> UpperLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> SpecLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> LowerLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class MeasureSeriesLimit
        {
            public string UpperText { get; set; } = string.Empty;
            public string SpecText { get; set; } = string.Empty;
            public string LowerText { get; set; } = string.Empty;
        }

        private sealed class CategoryBucket
        {
            public string Label { get; init; } = string.Empty;
            public List<DataRow> Rows { get; } = new();
        }

        private ProcessTrendFileInfo? _currentFile;
        private readonly List<string> _loadedFilePaths = new();
        private readonly List<PreviewColorChoice> _colorChoices = new();
        private readonly Dictionary<string, MeasureSeriesLimit> _seriesLimits = new(StringComparer.OrdinalIgnoreCase);
        private const string ValueColumnName = "Value";
        private const string DateColumnName = "Date";
        private bool _isRebuildingLimits;

        public event PropertyChangedEventHandler? PropertyChanged;

        public DailyDataTrendView()
        {
            InitializeComponent();
            DataContext = this;
            InitializeEditableColumnHeaderStyle();
            InitializeColorOptions();
            WireDataPreviewGrid();
        }

        private void WireDataPreviewGrid()
        {
            PreviewGraphViewBase.WirePreviewGrid(
                DataPreviewGrid,
                files => OpenSetupWindow(files),
                DeletePreviewRow,
                () => StatusText.Text = "Edited Data Preview cell value.",
                () =>
                {
                    if (_currentFile?.FullData == null)
                    {
                        return;
                    }

                    RowCountText.Text = _currentFile.FullData.Rows.Count.ToString("N0");
                    UpdateAxisSelectionOptions();
                    StatusText.Text = "Applied Data Preview row changes.";
                });
        }

        private void InitializeEditableColumnHeaderStyle()
        {
            DataPreviewGrid.ColumnHeaderStyle = BuildColumnHeaderStyle();
        }

        private Style BuildColumnHeaderStyle()
        {
            var style = new Style(typeof(DataGridColumnHeader));
            style.Setters.Add(new EventSetter(Control.MouseDoubleClickEvent, new MouseButtonEventHandler(DataGridColumnHeader_MouseDoubleClick)));

            var contextMenu = new ContextMenu();
            var renameItem = new MenuItem { Header = "Rename Column" };
            renameItem.Click += RenameColumnMenuItem_Click;
            contextMenu.Items.Add(renameItem);

            style.Setters.Add(new Setter(ContextMenuService.ContextMenuProperty, contextMenu));
            return style;
        }

        private void InitializeColorOptions()
        {
            PreviewGraphViewBase.InitializeDefaultColorOptions(PlotColorComboBox, _colorChoices);
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenSetupWindow();
        }

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0)
            {
                return;
            }

            OpenSetupWindow(files);
        }

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                StatusText.Foreground = new SolidColorBrush(Colors.Blue);
                StatusText.Text = "Drop files into the Data Preview grid.";
            }
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(102, 112, 133));
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private void OpenSetupWindow(IEnumerable<string>? initialFiles = null)
        {
            Window? owner = Window.GetWindow(this);
            var setupWindow = new DailyDataTrendSetupWindow(initialFiles)
            {
                Owner = owner
            };
            if (setupWindow.ShowDialog() != true || setupWindow.ResultFileInfo?.FullData == null)
            {
                return;
            }

            ApplyLoadedFileInfo(setupWindow.ResultFileInfo);
        }

        private void ApplyLoadedFileInfo(ProcessTrendFileInfo fileInfo)
        {
            _currentFile = fileInfo;
            _loadedFilePaths.Clear();
            foreach (string path in fileInfo.FilePath.Split('|', StringSplitOptions.RemoveEmptyEntries))
            {
                _loadedFilePaths.Add(path);
            }

            PreviewGraphViewBase.ApplyPreviewSummary(
                CurrentFileNameText,
                RowCountText,
                ColumnCountText,
                StatusText,
                string.IsNullOrWhiteSpace(fileInfo.Name)
                    ? string.Join(" + ", _loadedFilePaths.Select(Path.GetFileName))
                    : fileInfo.Name,
                fileInfo.FullData!.Rows.Count,
                fileInfo.FullData.Columns.Count,
                $"Loaded {_loadedFilePaths.Count:N0} file(s) and applied column settings.");
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
        }

        private void RefreshDataPreviewGrid()
        {
            DataPreviewGrid.ItemsSource = _currentFile?.FullData?.DefaultView;
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            _currentFile = null;
            _loadedFilePaths.Clear();
            _seriesLimits.Clear();
            PreviewGraphViewBase.ResetPreviewSummary(CurrentFileNameText, RowCountText, ColumnCountText, StatusText, "(No file)", "Load file and generate measure trend graph.");
            DataPreviewGrid.ItemsSource = null;
            LimitItemsPanel.Children.Clear();
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Load data first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var state = new DailyDataTrendReportState
            {
                Name = _currentFile.Name,
                FilePath = _currentFile.FilePath,
                Delimiter = _currentFile.Delimiter,
                HeaderRowNumber = _currentFile.HeaderRowNumber,
                SavedPlotColorIndex = PlotColorComboBox.SelectedIndex,
                RawData = GraphReportStorageHelper.CaptureRawTableData(_currentFile.FullData),
                UpperLimits = _seriesLimits.ToDictionary(pair => pair.Key, pair => pair.Value.UpperText, StringComparer.OrdinalIgnoreCase),
                SpecLimits = _seriesLimits.ToDictionary(pair => pair.Key, pair => pair.Value.SpecText, StringComparer.OrdinalIgnoreCase),
                LowerLimits = _seriesLimits.ToDictionary(pair => pair.Key, pair => pair.Value.LowerText, StringComparer.OrdinalIgnoreCase)
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "singlex-singley.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            DailyDataTrendReportState? state = GraphReportFileDialogHelper.LoadState<DailyDataTrendReportState>("Load Graph Report");
            if (state == null)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DataTable restoredTable = GraphReportStorageHelper.BuildTableFromRawData(state.RawData);
            var fileInfo = new ProcessTrendFileInfo
            {
                Name = state.Name,
                FilePath = state.FilePath,
                Delimiter = state.Delimiter,
                HeaderRowNumber = state.HeaderRowNumber,
                SavedPlotColorIndex = state.SavedPlotColorIndex,
                FullData = restoredTable
            };

            ApplyLoadedFileInfo(fileInfo);
            PlotColorComboBox.SelectedIndex = Math.Max(0, Math.Min(state.SavedPlotColorIndex, Math.Max(0, _colorChoices.Count - 1)));

            _seriesLimits.Clear();
            foreach (string key in state.UpperLimits.Keys
                         .Concat(state.SpecLimits.Keys)
                         .Concat(state.LowerLimits.Keys)
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                _seriesLimits[key] = new MeasureSeriesLimit
                {
                    UpperText = state.UpperLimits.TryGetValue(key, out string? upper) ? upper : string.Empty,
                    SpecText = state.SpecLimits.TryGetValue(key, out string? spec) ? spec : string.Empty,
                    LowerText = state.LowerLimits.TryGetValue(key, out string? lower) ? lower : string.Empty
                };
            }

            RebuildLimitEditor();
        }

        private void DeletePreviewRow()
        {
            if (_currentFile?.FullData == null)
            {
                return;
            }

            if (!DataPreviewGrid.TryDeleteSelectedRow(_currentFile.FullData))
            {
                return;
            }

            RowCountText.Text = _currentFile.FullData.Rows.Count.ToString("N0");
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
            StatusText.Text = "Selected row deleted from Data Preview.";
        }

        private void DataGridColumnHeader_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGridColumnHeader header)
            {
                return;
            }

            RenameColumn(header);
            e.Handled = true;
        }

        private void RenameColumnMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not MenuItem menuItem ||
                menuItem.Parent is not ContextMenu contextMenu ||
                contextMenu.PlacementTarget is not DataGridColumnHeader header)
            {
                return;
            }

            RenameColumn(header);
        }

        private void RenameColumn(DataGridColumnHeader header)
        {
            if (_currentFile?.FullData == null)
            {
                return;
            }

            string currentName = header.Column?.Header?.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(currentName) || !_currentFile.FullData.Columns.Contains(currentName))
            {
                return;
            }

            var renameWindow = new ColumnRenameWindow(_currentFile.FullData.Columns.Cast<DataColumn>().Select(column => column.ColumnName))
            {
                Owner = Window.GetWindow(this)
            };

            if (renameWindow.ShowDialog() != true)
            {
                return;
            }

            IReadOnlyDictionary<string, string> renamedColumns = renameWindow.RenamedColumns;
            if (!renamedColumns.TryGetValue(currentName, out string? newName) ||
                string.IsNullOrWhiteSpace(newName) ||
                string.Equals(currentName, newName, StringComparison.Ordinal))
            {
                if (!HasColumnRenameChanges(renamedColumns))
                {
                    return;
                }
            }

            ApplyColumnRenames(renamedColumns);

            ColumnCountText.Text = _currentFile.FullData.Columns.Count.ToString();
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
            StatusText.Text = HasColumnRenameChanges(renamedColumns)
                ? "Applied column rename changes."
                : "No column rename changes were applied.";
        }

        private bool HasColumnRenameChanges(IReadOnlyDictionary<string, string> renamedColumns)
        {
            return renamedColumns.Any(pair => !string.Equals(pair.Key, pair.Value, StringComparison.Ordinal));
        }

        private void ApplyColumnRenames(IReadOnlyDictionary<string, string> renamedColumns)
        {
            if (_currentFile?.FullData == null)
            {
                return;
            }

            Dictionary<string, string> pendingRenames = renamedColumns
                .Where(pair => !string.Equals(pair.Key, pair.Value, StringComparison.Ordinal))
                .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);

            if (pendingRenames.Count == 0)
            {
                return;
            }

            int renameIndex = 0;
            Dictionary<string, string> temporaryNames = new(StringComparer.OrdinalIgnoreCase);
            foreach ((string oldName, string newName) in pendingRenames)
            {
                RenameSeriesLimitKey(oldName, newName);
                RenameSpecLimitKey(_currentFile.UpperSpecLimits, oldName, newName);
                RenameSpecLimitKey(_currentFile.SpecLimits, oldName, newName);
                RenameSpecLimitKey(_currentFile.LowerSpecLimits, oldName, newName);

                string tempName;
                do
                {
                    tempName = $"__rename_{renameIndex++}_{Guid.NewGuid():N}";
                }
                while (_currentFile.FullData.Columns.Contains(tempName));

                _currentFile.FullData.Columns[oldName]!.ColumnName = tempName;
                temporaryNames[tempName] = newName;
            }

            foreach ((string tempName, string finalName) in temporaryNames)
            {
                _currentFile.FullData.Columns[tempName]!.ColumnName = finalName;
            }
        }

        private void RenameSeriesLimitKey(string oldName, string newName)
        {
            if (!_seriesLimits.TryGetValue(oldName, out MeasureSeriesLimit? limit))
            {
                return;
            }

            _seriesLimits.Remove(oldName);
            _seriesLimits[newName] = limit;
        }

        private static void RenameSpecLimitKey(IDictionary<string, double> limits, string oldName, string newName)
        {
            if (!limits.TryGetValue(oldName, out double value))
            {
                return;
            }

            limits.Remove(oldName);
            limits[newName] = value;
        }

        private void UpdateAxisSelectionOptions()
        {
            if (_currentFile?.FullData == null)
            {
                LimitItemsPanel.Children.Clear();
                return;
            }

            RebuildLimitEditor();
        }

        private void RebuildLimitEditor()
        {
            if (_isRebuildingLimits)
            {
                return;
            }

            _isRebuildingLimits = true;
            try
            {
                LimitItemsPanel.Children.Clear();
                if (_currentFile?.FullData == null || !_currentFile.FullData.Columns.Contains(ValueColumnName))
                {
                    LimitItemsPanel.Children.Add(new TextBlock
                    {
                        Text = "Load valid Date/Value data to enter limits.",
                        Foreground = new SolidColorBrush(Color.FromRgb(102, 112, 133))
                    });
                    return;
                }

                MeasureSeriesLimit limit = GetOrCreateLimit(ValueColumnName);
                LimitItemsPanel.Children.Add(BuildSeriesLimitEditor(ValueColumnName, limit));
            }
            finally
            {
                _isRebuildingLimits = false;
            }
        }

        private FrameworkElement BuildSeriesLimitEditor(string yName, MeasureSeriesLimit limit)
        {
            var group = new GroupBox
            {
                Header = yName,
                Margin = new Thickness(0, 0, 0, 8)
            };

            var grid = new Grid { Margin = new Thickness(8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(10) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var uslLabel = new TextBlock { Text = "USL", VerticalAlignment = System.Windows.VerticalAlignment.Center };
            Grid.SetColumn(uslLabel, 0);
            grid.Children.Add(uslLabel);

            var uslBox = new TextBox { Text = limit.UpperText, Tag = yName };
            uslBox.TextChanged += SeriesUpperLimitTextBox_TextChanged;
            Grid.SetColumn(uslBox, 2);
            grid.Children.Add(uslBox);

            var specLabel = new TextBlock { Text = "SPEC", VerticalAlignment = System.Windows.VerticalAlignment.Center };
            Grid.SetColumn(specLabel, 4);
            grid.Children.Add(specLabel);

            var specBox = new TextBox { Text = limit.SpecText, Tag = yName };
            specBox.TextChanged += SeriesSpecLimitTextBox_TextChanged;
            Grid.SetColumn(specBox, 6);
            grid.Children.Add(specBox);

            var lslLabel = new TextBlock { Text = "LSL", VerticalAlignment = System.Windows.VerticalAlignment.Center };
            Grid.SetColumn(lslLabel, 8);
            grid.Children.Add(lslLabel);

            var lslBox = new TextBox { Text = limit.LowerText, Tag = yName };
            lslBox.TextChanged += SeriesLowerLimitTextBox_TextChanged;
            Grid.SetColumn(lslBox, 10);
            grid.Children.Add(lslBox);

            group.Content = grid;
            return group;
        }

        private MeasureSeriesLimit GetOrCreateLimit(string seriesName)
        {
            if (!_seriesLimits.TryGetValue(seriesName, out MeasureSeriesLimit? limit))
            {
                limit = new MeasureSeriesLimit();
                _seriesLimits[seriesName] = limit;
            }

            return limit;
        }

        private void SeriesUpperLimitTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.Tag is string yName)
            {
                GetOrCreateLimit(yName).UpperText = textBox.Text.Trim();
            }
        }

        private void SeriesLowerLimitTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.Tag is string yName)
            {
                GetOrCreateLimit(yName).LowerText = textBox.Text.Trim();
            }
        }

        private void SeriesSpecLimitTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && textBox.Tag is string yName)
            {
                GetOrCreateLimit(yName).SpecText = textBox.Text.Trim();
            }
        }

        private void GenerateTrendButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Please load a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!ValidateFirstColumnDate(_currentFile.FullData, out string? errorMessage))
            {
                MessageBox.Show(errorMessage, "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!_currentFile.FullData.Columns.Contains(ValueColumnName))
            {
                MessageBox.Show("The loaded data does not contain a Value column.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PersistLimitsToFile();

            var buckets = BuildCategoryBuckets(_currentFile.FullData, new[] { DateColumnName });
            if (buckets.Count == 0)
            {
                MessageBox.Show("No valid rows were found for the Date column.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selectedColor = GetSelectedPlotColor();
            var labels = buckets.Select(bucket => bucket.Label).ToList();
            var model = CreateCategoryPlotModel("SingleX(Date) - SingleY", labels);
            var detailSb = new StringBuilder();
            detailSb.AppendLine($"X Axis: {DateColumnName}");
            detailSb.AppendLine($"Categories: {buckets.Count:N0}");
            detailSb.AppendLine();

            OxyColor color = GetColumnOxyColor(ValueColumnName, selectedColor);
            var values = BuildCategoryValues(buckets, ValueColumnName);
            var validPoints = values
                .Select((value, index) => new { value, index })
                .Where(item => !double.IsNaN(item.value))
                .Select(item => new DataPoint(item.index, item.value))
                .ToList();

            if (validPoints.Count == 0)
            {
                MessageBox.Show("No numeric Y-axis data could be plotted.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            model.Series.Add(BuildMeasureSeries(ValueColumnName, values, color));
            AddMeasureLimitLines(model, ValueColumnName, values.Count, color);

            var candidate = new ProcessTrendComputationCandidate
            {
                PairTitle = ValueColumnName,
                PlotModel = BuildSingleMeasureModel(ValueColumnName, labels, values, color),
                XAxisTitle = DateColumnName,
                YAxisTitle = ValueColumnName,
                RawPoints = validPoints
            };

            detailSb.AppendLine($"{ValueColumnName}: n={validPoints.Count:N0}, Avg={validPoints.Average(point => point.Y):F4}, Std={CalculateSampleStdDev(validPoints.Select(point => point.Y).ToList()):F4}");

            var graphDataList = new List<DailySamplingGraphData>
            {
                new DailySamplingGraphData
                {
                    FileName = CurrentFileNameText.Text,
                    Dates = _currentFile.FullData.Rows.Cast<DataRow>()
                        .Select(row => row[DateColumnName]?.ToString() ?? string.Empty)
                        .ToList(),
                    SampleNumbers = new List<string> { ValueColumnName },
                    Data = _currentFile.FullData,
                    UpperLimitValue = _currentFile.UpperSpecLimits.TryGetValue(ValueColumnName, out double usl) ? usl : null,
                    SpecValue = _currentFile.SpecLimits.TryGetValue(ValueColumnName, out double spec) ? spec : null,
                    LowerLimitValue = _currentFile.LowerSpecLimits.TryGetValue(ValueColumnName, out double lsl) ? lsl : null,
                    DataColor = selectedColor,
                    SpecColor = selectedColor,
                    UpperColor = selectedColor,
                    LowerColor = selectedColor,
                    IsXAxisDate = true,
                    XAxisMode = XAxisMode.Date,
                    DisplayMode = ValuePlotDisplayMode.Combined
                }
            };

            var resultWindow = new DailySamplingGraphViewerWindow(graphDataList, this)
            {
                WindowState = WindowState.Maximized
            };
            Window? ownerWindow = Window.GetWindow(this);
            if (ownerWindow is not null && ownerWindow.IsVisible)
            {
                resultWindow.Owner = ownerWindow;
            }
            resultWindow.Show();

            StatusText.Text = $"Displayed single-value graph across {buckets.Count:N0} date categories.";
        }

        private void PersistLimitsToFile()
        {
            if (_currentFile == null)
            {
                return;
            }

            _currentFile.UpperSpecLimits.Clear();
            _currentFile.SpecLimits.Clear();
            _currentFile.LowerSpecLimits.Clear();

            var limit = GetOrCreateLimit(ValueColumnName);
            if (GraphMakerParsingHelper.TryParseDouble(limit.UpperText, out double usl))
            {
                _currentFile.UpperSpecLimits[ValueColumnName] = usl;
            }

            if (GraphMakerParsingHelper.TryParseDouble(limit.SpecText, out double spec))
            {
                _currentFile.SpecLimits[ValueColumnName] = spec;
            }

            if (GraphMakerParsingHelper.TryParseDouble(limit.LowerText, out double lsl))
            {
                _currentFile.LowerSpecLimits[ValueColumnName] = lsl;
            }
        }

        private static List<CategoryBucket> BuildCategoryBuckets(DataTable table, IReadOnlyList<string> xColumns)
        {
            var buckets = new List<CategoryBucket>();
            var map = new Dictionary<string, CategoryBucket>(StringComparer.Ordinal);

            foreach (DataRow row in table.Rows)
            {
                string label = string.Join(" | ", xColumns.Select(column =>
                {
                    string text = row[column]?.ToString()?.Trim() ?? string.Empty;
                    return string.IsNullOrWhiteSpace(text) ? "(blank)" : text;
                }));

                if (!map.TryGetValue(label, out CategoryBucket? bucket))
                {
                    bucket = new CategoryBucket { Label = label };
                    map[label] = bucket;
                    buckets.Add(bucket);
                }

                bucket.Rows.Add(row);
            }

            return buckets;
        }

        private static List<double> BuildCategoryValues(IReadOnlyList<CategoryBucket> buckets, string yColumn)
        {
            var values = new List<double>(buckets.Count);
            foreach (CategoryBucket bucket in buckets)
            {
                var parsed = bucket.Rows.Select(row => TryGetRowDouble(row, yColumn))
                    .Where(value => value.HasValue)
                    .Select(value => value!.Value)
                    .ToList();

                values.Add(parsed.Count == 0 ? double.NaN : parsed.Average());
            }

            return values;
        }

        private static LineSeries BuildMeasureSeries(string title, IReadOnlyList<double> values, OxyColor color)
        {
            var series = new LineSeries
            {
                Title = title,
                Color = color,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3,
                MarkerFill = color,
                StrokeThickness = 2,
                RenderInLegend = true
            };

            for (int i = 0; i < values.Count; i++)
            {
                series.Points.Add(new DataPoint(i, values[i]));
            }

            return series;
        }

        private PlotModel BuildSingleMeasureModel(string title, IReadOnlyList<string> labels, IReadOnlyList<double> values, OxyColor color)
        {
            var model = CreateCategoryPlotModel(title, labels);
            model.Series.Add(BuildMeasureSeries(title, values, color));
            AddMeasureLimitLines(model, title, values.Count, color);
            return model;
        }

        private PlotModel CreateCategoryPlotModel(string title, IReadOnlyList<string> labels)
        {
            var model = new PlotModel { Title = title };
            var categoryAxis = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                Title = "Category",
                GapWidth = 0.5,
                Angle = labels.Count > 8 ? 35 : 0,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            };

            foreach (string label in labels)
            {
                categoryAxis.Labels.Add(label);
            }

            model.Axes.Add(categoryAxis);
            model.Axes.Add(new LinearAxis
            {
                Position = AxisPosition.Left,
                Title = "Value",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });
            return model;
        }

        private void AddMeasureLimitLines(PlotModel model, string yName, int categoryCount, OxyColor color)
        {
            if (_currentFile == null || categoryCount == 0)
            {
                return;
            }

            double start = -0.4;
            double end = Math.Max(0, categoryCount - 1) + 0.4;

            if (_currentFile.UpperSpecLimits.TryGetValue(yName, out double usl))
            {
                var line = new LineSeries
                {
                    Title = $"{yName} USL",
                    Color = color,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = true
                };
                line.Points.Add(new DataPoint(start, usl));
                line.Points.Add(new DataPoint(end, usl));
                model.Series.Add(line);
            }

            if (_currentFile.SpecLimits.TryGetValue(yName, out double spec))
            {
                var line = new LineSeries
                {
                    Title = $"{yName} SPEC",
                    Color = color,
                    LineStyle = LineStyle.Dot,
                    StrokeThickness = 1.8,
                    RenderInLegend = true
                };
                line.Points.Add(new DataPoint(start, spec));
                line.Points.Add(new DataPoint(end, spec));
                model.Series.Add(line);
            }

            if (_currentFile.LowerSpecLimits.TryGetValue(yName, out double lsl))
            {
                var line = new LineSeries
                {
                    Title = $"{yName} LSL",
                    Color = color,
                    LineStyle = LineStyle.Dash,
                    StrokeThickness = 1.8,
                    RenderInLegend = true
                };
                line.Points.Add(new DataPoint(start, lsl));
                line.Points.Add(new DataPoint(end, lsl));
                model.Series.Add(line);
            }
        }

        private Color GetSelectedPlotColor()
        {
            if (PlotColorComboBox.SelectedIndex >= 0 && PlotColorComboBox.SelectedIndex < _colorChoices.Count)
            {
                return _colorChoices[PlotColorComboBox.SelectedIndex].Color;
            }

            return Colors.Black;
        }

        private OxyColor GetColumnOxyColor(string columnName, Color fallbackColor)
        {
            if (_colorChoices.Count == 0)
            {
                return OxyColor.FromArgb(fallbackColor.A, fallbackColor.R, fallbackColor.G, fallbackColor.B);
            }

            int idx = Math.Abs(StringComparer.OrdinalIgnoreCase.GetHashCode(columnName)) % _colorChoices.Count;
            var color = _colorChoices[idx].Color;
            return OxyColor.FromArgb(color.A, color.R, color.G, color.B);
        }

        private static double? TryGetRowDouble(DataRow row, string columnName)
        {
            string text = row[columnName]?.ToString()?.Trim() ?? string.Empty;
            return GraphMakerParsingHelper.TryParseDouble(text, out double value) ? value : null;
        }

        private static bool ValidateFirstColumnDate(DataTable table, out string? errorMessage)
        {
            errorMessage = null;

            if (table.Columns.Count == 0)
            {
                errorMessage = "SingleX(Date) - SingleY requires at least one column, and the first column must contain date values.";
                return false;
            }

            string firstColumnName = table.Columns[0].ColumnName;
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                string text = table.Rows[rowIndex][0]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) ||
                    DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))
                {
                    continue;
                }

                errorMessage =
                    $"SingleX(Date) - SingleY requires the first column to contain date values such as mm/dd or yyyy-MM-dd. Invalid value '{text}' was found at row {rowIndex + 1:N0} in column '{firstColumnName}'.";
                return false;
            }

            return true;
        }

        private static double CalculateSampleStdDev(IReadOnlyList<double> values)
        {
            if (values.Count < 2)
            {
                return 0;
            }

            double avg = values.Average();
            double variance = values.Sum(value => Math.Pow(value - avg, 2)) / (values.Count - 1);
            return Math.Sqrt(variance);
        }

        public DateTime? ParseDateString(string dateStr)
        {
            return GraphMakerParsingHelper.TryParseDate(dateStr, out DateTime parsedDate)
                ? parsedDate
                : null;
        }
    }
}
