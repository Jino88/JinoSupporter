using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using Color = System.Windows.Media.Color;
using Colors = System.Windows.Media.Colors;
using JinoSupporter.Controls;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GraphMaker
{
    public partial class DailyDataTrendExtraView : GraphViewBase
    {
        private sealed class DailyDataTrendExtraReportState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public int SavedPlotColorIndex { get; set; }
            public bool ApplySameLimit { get; set; }
            public List<string> SelectedX { get; set; } = new();
            public List<string> SelectedY { get; set; } = new();
            public RawTableData RawData { get; set; } = new();
            public Dictionary<string, string> UpperLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> LowerLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class MeasureSeriesLimit
        {
            public string UpperText { get; set; } = string.Empty;
            public string LowerText { get; set; } = string.Empty;
        }

        private sealed class CategoryBucket
        {
            public string Label { get; init; } = string.Empty;
            public List<DataRow> Rows { get; } = new();
        }

        private ProcessTrendFileInfo? _currentFile;
        private readonly List<string> _loadedFilePaths = new();
        private readonly List<string> _pendingSetupFilePaths = new();
        private readonly List<PreviewColorChoice> _colorChoices = new();
        private readonly Dictionary<string, MeasureSeriesLimit> _seriesLimits = new(StringComparer.OrdinalIgnoreCase);
        private string _pendingSetupDelimiter = "\t";
        private int _pendingSetupHeaderRowNumber = 1;
        private bool _isRebuildingLimits;
        public DailyDataTrendExtraView()
        {
            InitializeComponent();
            DataContext = this;
            InitializeColorOptions();
            WireDataPreviewGrid();
            NotifyWebModuleSnapshotChanged();
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
                },
                RenameColumn);
        }

        private void InitializeColorOptions()
        {
            PreviewGraphViewBase.InitializeDefaultColorOptions(PlotColorComboBox, _colorChoices);
        }

        private void FileDropBox_FilesSelected(object sender, FilesSelectedEventArgs e)
        {
            OpenSetupWindow(e.FilePaths);
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select data file",
                Filter = "Text/CSV files (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                HandleWebDroppedFiles(dialog.FileNames);
            }
        }

        private void OpenSetupWindow(IEnumerable<string>? initialFiles = null)
        {
            Window? owner = Application.Current?.MainWindow ?? Window.GetWindow(this);
            var setupWindow = new DailyDataTrendSetupWindow(
                initialFiles,
                requireFirstColumnDate: false,
                convertWideToSingleY: false)
            {
                Owner = owner
            };
            if (setupWindow.ShowDialog() != true || setupWindow.ResultFileInfo?.FullData == null)
            {
                return;
            }

            ApplyLoadedFileInfo(setupWindow.ResultFileInfo);
        }

        public void HandleWebDroppedFiles(IReadOnlyList<string> filePaths)
        {
            string[] accepted = filePaths
                .Where(path =>
                {
                    string ext = Path.GetExtension(path).ToLowerInvariant();
                    return ext == ".txt" || ext == ".csv";
                })
                .ToArray();

            if (accepted.Length > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:MultiXMultiY] Accepted files: {string.Join(" | ", accepted)}");
                _pendingSetupFilePaths.Clear();
                _pendingSetupFilePaths.AddRange(accepted);
                _pendingSetupDelimiter = "\t";
                _pendingSetupHeaderRowNumber = 1;
                StatusText.Text = $"Dropped {accepted.Length:N0} file(s). Review settings and apply.";
                NotifyWebModuleSnapshotChanged();
                ApplyPendingWebSetup();
            }
        }

        private void ApplyPendingWebSetup()
        {
            if (_pendingSetupFilePaths.Count == 0)
            {
                return;
            }

            try
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:MultiXMultiY] ApplyPendingWebSetup delimiter={_pendingSetupDelimiter}, headerRow={_pendingSetupHeaderRowNumber}, files={string.Join(" | ", _pendingSetupFilePaths)}");
                ProcessTrendFileInfo fileInfo = DailyDataTrendSetupLoader.LoadProcessTrendFileInfo(
                    _pendingSetupFilePaths,
                    _pendingSetupDelimiter,
                    Math.Max(1, _pendingSetupHeaderRowNumber),
                    requireFirstColumnDate: false,
                    convertWideToSingleY: false);

                ApplyLoadedFileInfo(fileInfo);
                _pendingSetupFilePaths.Clear();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[GraphDrop:MultiXMultiY] ApplyPendingWebSetup failed: {ex}");
                StatusText.Text = ex.Message;
                NotifyWebModuleSnapshotChanged();
            }
        }

        private static string ParseWebDelimiter(string delimiter)
        {
            return delimiter switch
            {
                "comma" => ",",
                "space" => " ",
                _ => "\t"
            };
        }

        private void ReloadCurrentFilesWithDelimiter(string delimiter)
        {
            if (_loadedFilePaths.Count == 0 || _currentFile == null)
            {
                return;
            }

            List<string> selectedX = GetSelectedItems(XAxisListBox);
            List<string> selectedY = GetSelectedItems(YAxisListBox);
            Dictionary<string, MeasureSeriesLimit> savedLimits = _seriesLimits.ToDictionary(
                pair => pair.Key,
                pair => new MeasureSeriesLimit
                {
                    UpperText = pair.Value.UpperText,
                    LowerText = pair.Value.LowerText
                },
                StringComparer.OrdinalIgnoreCase);
            int selectedPlotColorIndex = PlotColorComboBox.SelectedIndex;
            bool applySameLimit = ApplySameLimitCheckBox.IsChecked == true;

            ProcessTrendFileInfo fileInfo = DailyDataTrendSetupLoader.LoadProcessTrendFileInfo(
                _loadedFilePaths,
                delimiter,
                Math.Max(1, _currentFile.HeaderRowNumber),
                requireFirstColumnDate: false,
                convertWideToSingleY: false);

            ApplyLoadedFileInfo(fileInfo);
            PlotColorComboBox.SelectedIndex = Math.Max(0, Math.Min(selectedPlotColorIndex, Math.Max(0, _colorChoices.Count - 1)));
            ApplySameLimitCheckBox.IsChecked = applySameLimit;

            _seriesLimits.Clear();
            foreach ((string key, MeasureSeriesLimit value) in savedLimits)
            {
                _seriesLimits[key] = value;
            }

            SetSelectedItems(XAxisListBox, selectedX);
            SetSelectedItems(YAxisListBox, selectedY);
            RebuildLimitEditor();
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
            NotifyWebModuleSnapshotChanged();
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
            if (!HasColumnRenameChanges(renamedColumns))
            {
                return;
            }

            ApplyColumnRenames(renamedColumns);
            ColumnCountText.Text = _currentFile.FullData.Columns.Count.ToString();
            RefreshDataPreviewGrid();
            UpdateAxisSelectionOptions();
            StatusText.Text = "Applied column rename changes.";
        }

        private static bool HasColumnRenameChanges(IReadOnlyDictionary<string, string> renamedColumns)
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
            XAxisListBox.ItemsSource = null;
            YAxisListBox.ItemsSource = null;
            LimitItemsPanel.Children.Clear();
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Load data first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var state = new DailyDataTrendExtraReportState
            {
                Name = _currentFile.Name,
                FilePath = _currentFile.FilePath,
                Delimiter = _currentFile.Delimiter,
                HeaderRowNumber = _currentFile.HeaderRowNumber,
                SavedPlotColorIndex = PlotColorComboBox.SelectedIndex,
                ApplySameLimit = ApplySameLimitCheckBox.IsChecked == true,
                SelectedX = GetSelectedItems(XAxisListBox),
                SelectedY = GetSelectedItems(YAxisListBox),
                RawData = GraphReportStorageHelper.CaptureRawTableData(_currentFile.FullData),
                UpperLimits = _seriesLimits.ToDictionary(pair => pair.Key, pair => pair.Value.UpperText, StringComparer.OrdinalIgnoreCase),
                LowerLimits = _seriesLimits.ToDictionary(pair => pair.Key, pair => pair.Value.LowerText, StringComparer.OrdinalIgnoreCase)
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "multix-multiy-trend.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            DailyDataTrendExtraReportState? state = GraphReportFileDialogHelper.LoadState<DailyDataTrendExtraReportState>("Load Graph Report");
            if (state == null)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var fileInfo = new ProcessTrendFileInfo
            {
                Name = state.Name,
                FilePath = state.FilePath,
                Delimiter = state.Delimiter,
                HeaderRowNumber = state.HeaderRowNumber,
                SavedPlotColorIndex = state.SavedPlotColorIndex,
                FullData = GraphReportStorageHelper.BuildTableFromRawData(state.RawData)
            };

            ApplyLoadedFileInfo(fileInfo);
            PlotColorComboBox.SelectedIndex = Math.Max(0, Math.Min(state.SavedPlotColorIndex, Math.Max(0, _colorChoices.Count - 1)));
            ApplySameLimitCheckBox.IsChecked = state.ApplySameLimit;

            _seriesLimits.Clear();
            foreach (string key in state.UpperLimits.Keys.Concat(state.LowerLimits.Keys).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                _seriesLimits[key] = new MeasureSeriesLimit
                {
                    UpperText = state.UpperLimits.TryGetValue(key, out string? upper) ? upper : string.Empty,
                    LowerText = state.LowerLimits.TryGetValue(key, out string? lower) ? lower : string.Empty
                };
            }

            SetSelectedItems(XAxisListBox, state.SelectedX);
            SetSelectedItems(YAxisListBox, state.SelectedY);
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

        private void UpdateAxisSelectionOptions()
        {
            var previousX = GetSelectedItems(XAxisListBox);
            var previousY = GetSelectedItems(YAxisListBox);

            XAxisListBox.ItemsSource = null;
            YAxisListBox.ItemsSource = null;

            if (_currentFile?.FullData == null)
            {
                LimitItemsPanel.Children.Clear();
                return;
            }

            var allColumns = _currentFile.FullData.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList();
            var numericColumns = allColumns.Where(column => IsLikelyNumericColumn(_currentFile.FullData, column)).ToList();

            XAxisListBox.ItemsSource = allColumns;
            YAxisListBox.ItemsSource = numericColumns;

            if (previousX.Count == 0 && previousY.Count == 0)
            {
                SetSelectedItems(XAxisListBox, allColumns.Except(numericColumns, StringComparer.OrdinalIgnoreCase));
                SetSelectedItems(YAxisListBox, numericColumns);
            }
            else
            {
                SetSelectedItems(XAxisListBox, previousX);
                SetSelectedItems(YAxisListBox, previousY);
            }

            RebuildLimitEditor();
        }

        private static List<string> GetSelectedItems(ListBox listBox)
        {
            return listBox.SelectedItems.Cast<object>()
                .Select(item => item.ToString() ?? string.Empty)
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .ToList();
        }

        private static void SetSelectedItems(ListBox listBox, IEnumerable<string> values)
        {
            var wanted = new HashSet<string>(values, StringComparer.OrdinalIgnoreCase);
            listBox.SelectedItems.Clear();
            foreach (var item in listBox.Items)
            {
                if (wanted.Contains(item?.ToString() ?? string.Empty))
                {
                    listBox.SelectedItems.Add(item);
                }
            }
        }

        private void AxisSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            RebuildLimitEditor();
        }

        private void ApplySameLimitCheckBox_Changed(object sender, RoutedEventArgs e)
        {
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
                var selectedY = GetSelectedItems(YAxisListBox);
                if (selectedY.Count == 0)
                {
                    LimitItemsPanel.Children.Add(new TextBlock
                    {
                        Text = "Select Y-axis series to enter limits.",
                        Foreground = new SolidColorBrush(Color.FromRgb(102, 112, 133))
                    });
                    return;
                }

                if (ApplySameLimitCheckBox.IsChecked == true)
                {
                    string sharedUpper = selectedY.Select(GetOrCreateLimit).Select(limit => limit.UpperText).FirstOrDefault(text => !string.IsNullOrWhiteSpace(text)) ?? string.Empty;
                    string sharedLower = selectedY.Select(GetOrCreateLimit).Select(limit => limit.LowerText).FirstOrDefault(text => !string.IsNullOrWhiteSpace(text)) ?? string.Empty;

                    LimitItemsPanel.Children.Add(BuildSharedLimitEditor("USL", sharedUpper, SharedUpperLimitTextBox_TextChanged));
                    LimitItemsPanel.Children.Add(BuildSharedLimitEditor("LSL", sharedLower, SharedLowerLimitTextBox_TextChanged));
                    return;
                }

                foreach (string yName in selectedY)
                {
                    MeasureSeriesLimit limit = GetOrCreateLimit(yName);
                    LimitItemsPanel.Children.Add(BuildSeriesLimitEditor(yName, limit));
                }
            }
            finally
            {
                _isRebuildingLimits = false;
            }
        }

        private FrameworkElement BuildSharedLimitEditor(string label, string value, TextChangedEventHandler changed)
        {
            var panel = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };
            panel.Children.Add(new TextBlock
            {
                Text = $"{label} (All selected Y)",
                FontWeight = System.Windows.FontWeights.SemiBold
            });

            var textBox = new TextBox
            {
                Margin = new Thickness(0, 4, 0, 0),
                Text = value
            };
            textBox.TextChanged += changed;
            panel.Children.Add(textBox);
            return panel;
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

            var uslLabel = new TextBlock { Text = "USL", VerticalAlignment = System.Windows.VerticalAlignment.Center };
            Grid.SetColumn(uslLabel, 0);
            grid.Children.Add(uslLabel);

            var uslBox = new TextBox { Text = limit.UpperText, Tag = yName };
            uslBox.TextChanged += SeriesUpperLimitTextBox_TextChanged;
            Grid.SetColumn(uslBox, 2);
            grid.Children.Add(uslBox);

            var lslLabel = new TextBlock { Text = "LSL", VerticalAlignment = System.Windows.VerticalAlignment.Center };
            Grid.SetColumn(lslLabel, 4);
            grid.Children.Add(lslLabel);

            var lslBox = new TextBox { Text = limit.LowerText, Tag = yName };
            lslBox.TextChanged += SeriesLowerLimitTextBox_TextChanged;
            Grid.SetColumn(lslBox, 6);
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

        private void SharedUpperLimitTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox textBox)
            {
                return;
            }

            foreach (string yName in GetSelectedItems(YAxisListBox))
            {
                GetOrCreateLimit(yName).UpperText = textBox.Text.Trim();
            }
        }

        private void SharedLowerLimitTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox textBox)
            {
                return;
            }

            foreach (string yName in GetSelectedItems(YAxisListBox))
            {
                GetOrCreateLimit(yName).LowerText = textBox.Text.Trim();
            }
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

        private void GenerateTrendButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Please load a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var selectedX = GetSelectedItems(XAxisListBox);
            var selectedY = GetSelectedItems(YAxisListBox);
            if (selectedX.Count == 0)
            {
                MessageBox.Show("Select at least one X-axis column.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (selectedY.Count == 0)
            {
                MessageBox.Show("Select at least one Y-axis column.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PersistLimitsToFile();

            var selectedColor = GetSelectedPlotColor();
            var detailSb = new StringBuilder();
            detailSb.AppendLine($"X Axis: {string.Join(", ", selectedX)}");
            var pairResults = new List<ProcessPairPlotResult>();

            foreach (string xName in selectedX)
            {
                var rowItems = BuildRowItems(_currentFile.FullData, xName);
                if (rowItems.Count == 0)
                {
                    continue;
                }

                bool useNumericXAxis = IsLikelyNumericColumn(_currentFile.FullData, xName);
                var labels = rowItems
                    .OrderBy(item => item.CategoryIndex)
                    .Select(item => item.Label)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var model = useNumericXAxis
                    ? CreateNumericPlotModel($"X: {xName}", xName)
                    : CreateCategoryPlotModel($"X: {xName}", xName, labels);
                var candidates = new List<ProcessTrendComputationCandidate>();
                var xDetailSb = new StringBuilder();
                xDetailSb.AppendLine($"X Axis: {xName}");
                xDetailSb.AppendLine($"Rows: {rowItems.Count:N0}");
                xDetailSb.AppendLine();

                foreach (string yName in selectedY)
                {
                    OxyColor color = GetColumnOxyColor(yName, selectedColor);
                    var values = BuildRowValues(rowItems, yName);
                    var validPoints = values
                        .Select((value, index) => new { value, index })
                        .Where(item => !double.IsNaN(item.value))
                        .Select(item => new DataPoint(
                            useNumericXAxis ? (rowItems[item.index].NumericX ?? item.index) : rowItems[item.index].CategoryIndex,
                            item.value))
                        .ToList();

                    if (validPoints.Count == 0)
                    {
                        xDetailSb.AppendLine($"{yName}: no numeric data");
                        continue;
                    }

                    if (useNumericXAxis)
                    {
                        AddSeriesAreaHull(model, yName, validPoints, color);
                    }

                    model.Series.Add(BuildMeasureSeries(yName, rowItems, values, color, useNumericXAxis));
                    AddMeasureLimitLines(model, yName, rowItems, color, useNumericXAxis);

                    candidates.Add(new ProcessTrendComputationCandidate
                    {
                        PairTitle = yName,
                        PlotModel = BuildSingleMeasureModel(yName, xName, rowItems, labels, values, color, useNumericXAxis),
                        XAxisTitle = xName,
                        YAxisTitle = yName,
                        RawPoints = validPoints
                    });

                    xDetailSb.AppendLine($"{yName}: n={validPoints.Count:N0}, Avg={validPoints.Average(point => point.Y):F4}, Std={CalculateSampleStdDev(validPoints.Select(point => point.Y).ToList()):F4}");
                }

                if (candidates.Count == 0)
                {
                    continue;
                }

                pairResults.Add(new ProcessPairPlotResult
                {
                    PairTitle = xName,
                    PlotModel = model,
                    DetailText = xDetailSb.ToString().TrimEnd(),
                    XAxisTitle = xName,
                    YAxisTitle = string.Join(", ", selectedY),
                    RawPoints = candidates[0].RawPoints,
                    ComputationCandidates = candidates
                });
            }

            if (pairResults.Count == 0)
            {
                MessageBox.Show("No numeric Y-axis data could be plotted.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var resultWindow = new ProcessFlowTrendResultWindow(pairResults, "MultiX - MultiY Trend Result");
            Window? ownerWindow = Window.GetWindow(this);
            if (ownerWindow is not null && ownerWindow.IsVisible)
            {
                resultWindow.Owner = ownerWindow;
            }
            resultWindow.Show();

            StatusText.Text = $"Displayed {pairResults.Count:N0} X-axis graph(s) with {selectedY.Count:N0} Y series.";
        }

        private void PersistLimitsToFile()
        {
            if (_currentFile == null)
            {
                return;
            }

            _currentFile.UpperSpecLimits.Clear();
            _currentFile.LowerSpecLimits.Clear();

            foreach (string yName in GetSelectedItems(YAxisListBox))
            {
                var limit = GetOrCreateLimit(yName);
                if (GraphMakerParsingHelper.TryParseDouble(limit.UpperText, out double usl))
                {
                    _currentFile.UpperSpecLimits[yName] = usl;
                }

                if (GraphMakerParsingHelper.TryParseDouble(limit.LowerText, out double lsl))
                {
                    _currentFile.LowerSpecLimits[yName] = lsl;
                }
            }
        }

        private sealed class RowItem
        {
            public string Label { get; init; } = string.Empty;
            public double? NumericX { get; init; }
            public int CategoryIndex { get; init; }
            public DataRow Row { get; init; } = null!;
        }

        private static List<RowItem> BuildRowItems(DataTable table, string xColumn)
        {
            var items = new List<RowItem>();
            var categoryIndexByLabel = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in table.Rows)
            {
                string text = row[xColumn]?.ToString()?.Trim() ?? string.Empty;
                string label = string.IsNullOrWhiteSpace(text) ? "(blank)" : text;
                if (!categoryIndexByLabel.TryGetValue(label, out int categoryIndex))
                {
                    categoryIndex = categoryIndexByLabel.Count;
                    categoryIndexByLabel[label] = categoryIndex;
                }

                items.Add(new RowItem
                {
                    Label = label,
                    NumericX = GraphMakerParsingHelper.TryParseDouble(text, out double numericX) ? numericX : null,
                    CategoryIndex = categoryIndex,
                    Row = row
                });
            }

            return items;
        }

        private static List<double> BuildRowValues(IReadOnlyList<RowItem> rowItems, string yColumn)
        {
            var values = new List<double>(rowItems.Count);
            foreach (RowItem rowItem in rowItems)
            {
                values.Add(TryGetRowDouble(rowItem.Row, yColumn) ?? double.NaN);
            }

            return values;
        }

        private static Series BuildMeasureSeries(string title, IReadOnlyList<RowItem> rowItems, IReadOnlyList<double> values, OxyColor color, bool useNumericXAxis)
        {
            if (useNumericXAxis)
            {
                var scatterSeries = new ScatterSeries
                {
                    Title = title,
                    MarkerType = MarkerType.Circle,
                    MarkerSize = 3.5,
                    MarkerFill = color,
                    RenderInLegend = true
                };

                for (int i = 0; i < values.Count; i++)
                {
                    if (double.IsNaN(values[i]))
                    {
                        continue;
                    }

                    double x = rowItems[i].NumericX ?? i;
                    scatterSeries.Points.Add(new ScatterPoint(x, values[i]));
                }

                return scatterSeries;
            }

            var categorySeries = new LineSeries
            {
                Title = title,
                Color = color,
                MarkerType = MarkerType.Circle,
                MarkerSize = 3.5,
                MarkerFill = color,
                MarkerStroke = color,
                StrokeThickness = 0,
                LineStyle = LineStyle.None,
                RenderInLegend = true
            };

            for (int i = 0; i < values.Count; i++)
            {
                if (double.IsNaN(values[i]))
                {
                    continue;
                }

                categorySeries.Points.Add(new DataPoint(rowItems[i].CategoryIndex, values[i]));
            }

            return categorySeries;
        }

        private PlotModel BuildSingleMeasureModel(
            string title,
            string xAxisTitle,
            IReadOnlyList<RowItem> rowItems,
            IReadOnlyList<string> labels,
            IReadOnlyList<double> values,
            OxyColor color,
            bool useNumericXAxis)
        {
            var model = useNumericXAxis
                ? CreateNumericPlotModel(title, xAxisTitle)
                : CreateCategoryPlotModel(title, xAxisTitle, labels);
            model.Series.Add(BuildMeasureSeries(title, rowItems, values, color, useNumericXAxis));
            AddMeasureLimitLines(model, title, rowItems, color, useNumericXAxis);
            return model;
        }

        private PlotModel CreateCategoryPlotModel(string title, string xAxisTitle, IReadOnlyList<string> labels)
        {
            var model = new PlotModel { Title = title };
            var categoryAxis = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                Title = xAxisTitle,
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

        private PlotModel CreateNumericPlotModel(string title, string xAxisTitle)
        {
            var model = new PlotModel { Title = title };
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
                Title = "Value",
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.Dot
            });
            return model;
        }

        private static void AddSeriesAreaHull(PlotModel model, string title, IReadOnlyList<DataPoint> points, OxyColor color)
        {
            List<DataPoint> hull = BuildConvexHull(points);
            if (hull.Count < 3)
            {
                return;
            }

            var polygon = new PolygonAnnotation
            {
                Layer = AnnotationLayer.BelowSeries,
                Fill = OxyColor.FromAColor(60, color),
                Stroke = color,
                StrokeThickness = 1.2,
                LineStyle = LineStyle.Solid,
                ToolTip = $"{title} Area"
            };

            foreach (DataPoint point in hull)
            {
                polygon.Points.Add(new DataPoint(point.X, point.Y));
            }

            model.Annotations.Add(polygon);
        }

        private static List<DataPoint> BuildConvexHull(IReadOnlyList<DataPoint> points)
        {
            List<DataPoint> uniquePoints = points
                .GroupBy(point => (Math.Round(point.X, 8), Math.Round(point.Y, 8)))
                .Select(group => group.First())
                .OrderBy(point => point.X)
                .ThenBy(point => point.Y)
                .ToList();

            if (uniquePoints.Count < 3)
            {
                return uniquePoints;
            }

            var lower = new List<DataPoint>();
            foreach (DataPoint point in uniquePoints)
            {
                while (lower.Count >= 2 && Cross(lower[^2], lower[^1], point) <= 0)
                {
                    lower.RemoveAt(lower.Count - 1);
                }

                lower.Add(point);
            }

            var upper = new List<DataPoint>();
            for (int i = uniquePoints.Count - 1; i >= 0; i--)
            {
                DataPoint point = uniquePoints[i];
                while (upper.Count >= 2 && Cross(upper[^2], upper[^1], point) <= 0)
                {
                    upper.RemoveAt(upper.Count - 1);
                }

                upper.Add(point);
            }

            lower.RemoveAt(lower.Count - 1);
            upper.RemoveAt(upper.Count - 1);
            lower.AddRange(upper);
            return lower;
        }

        private static double Cross(DataPoint origin, DataPoint a, DataPoint b)
        {
            return (a.X - origin.X) * (b.Y - origin.Y) - (a.Y - origin.Y) * (b.X - origin.X);
        }

        private void AddMeasureLimitLines(PlotModel model, string yName, IReadOnlyList<RowItem> rowItems, OxyColor color, bool useNumericXAxis)
        {
            if (_currentFile == null || rowItems.Count == 0)
            {
                return;
            }

            double start = -0.4;
            double end = Math.Max(0, rowItems.Max(item => item.CategoryIndex)) + 0.4;
            if (useNumericXAxis)
            {
                List<double> xs = rowItems.Where(item => item.NumericX.HasValue).Select(item => item.NumericX!.Value).ToList();
                if (xs.Count > 0)
                {
                    start = xs.Min();
                    end = xs.Max();
                }
            }

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

        private static bool IsLikelyNumericColumn(DataTable table, string columnName)
        {
            int checkedCount = 0;
            int numericCount = 0;

            foreach (DataRow row in table.Rows)
            {
                string text = row[columnName]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                checkedCount++;
                if (GraphMakerParsingHelper.TryParseDouble(text, out _))
                {
                    numericCount++;
                }

                if (checkedCount >= 20)
                {
                    break;
                }
            }

            return checkedCount > 0 && numericCount == checkedCount;
        }

        private static double? TryGetRowDouble(DataRow row, string columnName)
        {
            string text = row[columnName]?.ToString()?.Trim() ?? string.Empty;
            return GraphMakerParsingHelper.TryParseDouble(text, out double value) ? value : null;
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

        public object GetWebModuleSnapshot()
        {
            return new
            {
                moduleType = "GraphMakerDailyDataTrendExtra",
                fileName = CurrentFileNameText.Text ?? "(No file)",
                rowCount = RowCountText.Text ?? "0",
                columnCount = ColumnCountText.Text ?? "0",
                status = StatusText.Text ?? string.Empty,
                pendingSetup = _pendingSetupFilePaths.Count == 0 ? null : new
                {
                    files = _pendingSetupFilePaths.Select(Path.GetFileName).ToArray(),
                    filePaths = _pendingSetupFilePaths.ToArray(),
                    delimiter = _pendingSetupDelimiter == "\t" ? "tab" : _pendingSetupDelimiter == "," ? "comma" : "space",
                    headerRow = _pendingSetupHeaderRowNumber.ToString(CultureInfo.InvariantCulture)
                },
                delimiter = _currentFile == null ? "tab" : _currentFile.Delimiter == "\t" ? "tab" : _currentFile.Delimiter == "," ? "comma" : "space",
                plotColor = PlotColorComboBox.SelectedItem is PreviewColorChoice selectedColorChoice
                    ? selectedColorChoice.Name
                    : string.Empty,
                plotColorOptions = _colorChoices.Select(choice => new
                {
                    name = choice.Name,
                    hex = $"#{choice.Color.R:X2}{choice.Color.G:X2}{choice.Color.B:X2}"
                }).ToArray(),
                applySameLimit = ApplySameLimitCheckBox.IsChecked == true,
                selectedX = XAxisListBox.SelectedItems.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                selectedY = YAxisListBox.SelectedItems.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                xAxisOptions = XAxisListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                yAxisOptions = YAxisListBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                previewColumns = BuildPreviewColumns(_currentFile?.FullData),
                previewRows = BuildPreviewRows(_currentFile?.FullData),
                limits = _seriesLimits.OrderBy(pair => pair.Key).Select(pair => new
                {
                    name = pair.Key,
                    upper = pair.Value.UpperText,
                    lower = pair.Value.LowerText
                }).ToArray()
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("pendingDelimiter", out JsonElement pendingDelimiterElement))
            {
                string pendingDelimiter = pendingDelimiterElement.GetString() ?? "tab";
                _pendingSetupDelimiter = ParseWebDelimiter(pendingDelimiter);
            }

            if (payload.TryGetProperty("pendingHeaderRow", out JsonElement pendingHeaderRowElement) &&
                int.TryParse(pendingHeaderRowElement.GetString(), out int pendingHeaderRowNumber) &&
                pendingHeaderRowNumber > 0)
            {
                _pendingSetupHeaderRowNumber = pendingHeaderRowNumber;
            }

            if (payload.TryGetProperty("plotColor", out JsonElement plotColorElement))
            {
                string plotColor = plotColorElement.GetString() ?? string.Empty;
                PreviewColorChoice? match = _colorChoices.FirstOrDefault(item => string.Equals(item.Name, plotColor, StringComparison.Ordinal));
                if (match is not null)
                {
                    PlotColorComboBox.SelectedItem = match;
                }
            }

            if (payload.TryGetProperty("delimiter", out JsonElement delimiterElement))
            {
                string webDelimiter = delimiterElement.GetString() ?? "tab";
                string parsedDelimiter = ParseWebDelimiter(webDelimiter);
                if (_currentFile == null || !string.Equals(_currentFile.Delimiter, parsedDelimiter, StringComparison.Ordinal))
                {
                    ReloadCurrentFilesWithDelimiter(parsedDelimiter);
                }
            }

            if (payload.TryGetProperty("applySameLimit", out JsonElement applySameLimitElement))
            {
                ApplySameLimitCheckBox.IsChecked = applySameLimitElement.GetBoolean();
            }

            if (payload.TryGetProperty("selectedX", out JsonElement selectedXElement) &&
                selectedXElement.ValueKind == JsonValueKind.Array)
            {
                SetSelectedItems(XAxisListBox, selectedXElement.EnumerateArray().Select(item => item.GetString() ?? string.Empty));
            }

            if (payload.TryGetProperty("selectedY", out JsonElement selectedYElement) &&
                selectedYElement.ValueKind == JsonValueKind.Array)
            {
                SetSelectedItems(YAxisListBox, selectedYElement.EnumerateArray().Select(item => item.GetString() ?? string.Empty));
                RebuildLimitEditor();
            }

            if (payload.TryGetProperty("limits", out JsonElement limitsElement) &&
                limitsElement.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement item in limitsElement.EnumerateArray())
                {
                    string name = item.TryGetProperty("name", out JsonElement nameElement) ? nameElement.GetString() ?? string.Empty : string.Empty;
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    MeasureSeriesLimit limit = GetOrCreateLimit(name);
                    if (item.TryGetProperty("upper", out JsonElement upperElement))
                    {
                        limit.UpperText = upperElement.GetString() ?? string.Empty;
                    }

                    if (item.TryGetProperty("lower", out JsonElement lowerElement))
                    {
                        limit.LowerText = lowerElement.GetString() ?? string.Empty;
                    }
                }

                RebuildLimitEditor();
            }

            NotifyWebModuleSnapshotChanged();
            return GetWebModuleSnapshot();
        }

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "browse-file":
                    BrowseButton_Click(this, new RoutedEventArgs());
                    break;
                case "load-report":
                    LoadReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "save-report":
                    SaveReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "remove-file":
                    RemoveFileButton_Click(this, new RoutedEventArgs());
                    break;
                case "generate-trend":
                    GenerateTrendButton_Click(this, new RoutedEventArgs());
                    break;
                case "apply-inline-setup":
                    ApplyPendingWebSetup();
                    break;
                case "cancel-inline-setup":
                    _pendingSetupFilePaths.Clear();
                    NotifyWebModuleSnapshotChanged();
                    break;
            }

            return GetWebModuleSnapshot();
        }

        private static string[] BuildPreviewColumns(DataTable? table)
        {
            return table?.Columns.Cast<DataColumn>().Take(24).Select(column => column.ColumnName ?? string.Empty).ToArray()
                ?? Array.Empty<string>();
        }

        private static string[][] BuildPreviewRows(DataTable? table)
        {
            if (table == null)
            {
                return Array.Empty<string[]>();
            }

            DataColumn[] columns = table.Columns.Cast<DataColumn>().Take(24).ToArray();
            return table.Rows.Cast<DataRow>()
                .Take(40)
                .Select(row => columns.Select(column => row[column]?.ToString() ?? string.Empty).ToArray())
                .ToArray();
        }

    }
}
