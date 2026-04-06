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
using System.Windows.Media;
using OxyPlot;
using OxyPlot.Annotations;
using OxyPlot.Axes;
using OxyPlot.Series;
using UserControl = System.Windows.Controls.UserControl;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace GraphMaker
{
    public class HeatMapFileInfo
    {
        public string Name { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public DataTable? FullData { get; set; }
        public List<string>? HeaderRow { get; set; }
        public string Delimiter { get; set; } = "\t";
        public int HeaderRowNumber { get; set; } = 1;
        public int SavedCondition1Index { get; set; } = -1;
        public int SavedCondition2Index { get; set; } = -1;
        public int SavedExtraCondition1Index { get; set; } = -1;
        public int SavedExtraCondition2Index { get; set; } = -1;
        public int SavedResultIndex { get; set; } = -1;
        public int SavedNgColumnIndex { get; set; } = -1;
        public int SavedInputColumnIndex { get; set; } = -1;
        public bool UseFormulaResult { get; set; }
    }

    public partial class HeatMapView : UserControl, INotifyPropertyChanged
    {
        private sealed class HeatMapReportState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public RawTableData RawData { get; set; } = new();
            public int SavedCondition1Index { get; set; } = -1;
            public int SavedCondition2Index { get; set; } = -1;
            public int SavedExtraCondition1Index { get; set; } = -1;
            public int SavedExtraCondition2Index { get; set; } = -1;
            public int SavedResultIndex { get; set; } = -1;
            public int SavedNgColumnIndex { get; set; } = -1;
            public int SavedInputColumnIndex { get; set; } = -1;
            public bool UseFormulaResult { get; set; }
        }

        private HeatMapFileInfo? _currentFile;
        private PlotModel? _heatMapModel;

        public PlotModel? HeatMapModel
        {
            get => _heatMapModel;
            set
            {
                _heatMapModel = value;
                OnPropertyChanged(nameof(HeatMapModel));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public event Action? WebModuleSnapshotChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public HeatMapView()
        {
            InitializeComponent();
            DataContext = this;
            HeatMapModel = GraphSampleModelFactory.CreateHeatMapSample();
            NotifyWebModuleSnapshotChanged();
        }

        #region Drag and Drop

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            var firstTxtFile = files.FirstOrDefault(f =>
                string.Equals(Path.GetExtension(f), ".txt", StringComparison.OrdinalIgnoreCase));

            if (firstTxtFile == null)
            {
                MessageBox.Show("Only TXT files are supported.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            LoadFile(firstTxtFile);
        }

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                SetDropHintForeground(sender, Colors.Blue);
            }
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            SetDropHintForeground(sender, System.Windows.Media.Color.FromRgb(51, 65, 85));
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private static void SetDropHintForeground(object sender, System.Windows.Media.Color color)
        {
            if (sender is Border border &&
                border.Child is Panel panel &&
                panel.Children.OfType<TextBlock>().FirstOrDefault() is TextBlock textBlock)
            {
                textBlock.Foreground = new SolidColorBrush(color);
            }
        }

        #endregion

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Select TXT File",
                Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LoadFile(openFileDialog.FileName);
            }
        }

        private void LoadFile(string filePath)
        {
            try
            {
                if (!string.Equals(Path.GetExtension(filePath), ".txt", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Only TXT files are supported.", "Notice",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (_currentFile != null &&
                    string.Equals(_currentFile.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show($"{Path.GetFileName(filePath)} is already selected.",
                        "Duplicate", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _currentFile = new HeatMapFileInfo
                {
                    Name = Path.GetFileName(filePath),
                    FilePath = filePath
                };

                LoadFileData(_currentFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadFileData(HeatMapFileInfo fileInfo)
        {
            try
            {
                if (fileInfo.FullData != null && fileInfo.FullData.Columns.Count > 0)
                {
                    fileInfo.HeaderRow = fileInfo.FullData.Columns.Cast<DataColumn>()
                        .Select(column => column.ColumnName)
                        .ToList();

                    PreviewGraphViewBase.ApplyPreviewSummary(
                        CurrentFileNameText,
                        RowCountText,
                        ColumnCountText,
                        StatusText,
                        fileInfo.Name,
                        fileInfo.FullData.Rows.Count,
                        fileInfo.FullData.Columns.Count,
                        $"{fileInfo.Name} loaded from report. Configure Condition1/Condition2/(Optional)Extra1/Extra2/Result column.");
                    DataPreviewGrid.ItemsSource = fileInfo.FullData.DefaultView;
                    UpdateColumnSelectors(fileInfo);
                    return;
                }

                var lines = File.ReadAllLines(fileInfo.FilePath);
                if (lines.Length == 0)
                {
                    MessageBox.Show("File is empty.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var headerRowIndex = fileInfo.HeaderRowNumber - 1;
                if (headerRowIndex < 0 || headerRowIndex >= lines.Length)
                {
                    MessageBox.Show("Header row is out of file range.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var rawHeader = GraphMakerTableHelper.SplitLine(lines[headerRowIndex], fileInfo.Delimiter);
                if (rawHeader.Length == 0)
                {
                    MessageBox.Show("Unable to read header. Check delimiter/header row.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var headers = GraphMakerTableHelper.BuildUniqueHeaders(rawHeader);
                var dataTable = new DataTable();
                foreach (var header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                for (var i = headerRowIndex + 1; i < lines.Length; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                    {
                        continue;
                    }

                    var values = GraphMakerTableHelper.SplitLine(lines[i], fileInfo.Delimiter);
                    if (values.Length == 0)
                    {
                        continue;
                    }

                    var row = dataTable.NewRow();
                    for (var j = 0; j < Math.Min(values.Length, headers.Count); j++)
                    {
                        row[j] = values[j];
                    }

                    dataTable.Rows.Add(row);
                }

                fileInfo.HeaderRow = headers;
                fileInfo.FullData = dataTable;

                PreviewGraphViewBase.ApplyPreviewSummary(
                    CurrentFileNameText,
                    RowCountText,
                    ColumnCountText,
                    StatusText,
                    fileInfo.Name,
                    dataTable.Rows.Count,
                    dataTable.Columns.Count,
                    $"{fileInfo.Name} loaded. Configure Condition1/Condition2/(Optional)Extra1/Extra2/Result column.");
                DataPreviewGrid.ItemsSource = dataTable.DefaultView;

                UpdateColumnSelectors(fileInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"   Error: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateColumnSelectors(HeatMapFileInfo fileInfo)
        {
            Condition1ComboBox.Items.Clear();
            Condition2ComboBox.Items.Clear();
            ExtraCondition1ComboBox.Items.Clear();
            ExtraCondition2ComboBox.Items.Clear();
            ResultColumnComboBox.Items.Clear();
            NgColumnComboBox.Items.Clear();
            InputColumnComboBox.Items.Clear();

            if (fileInfo.HeaderRow == null || fileInfo.HeaderRow.Count == 0)
            {
                return;
            }

            ExtraCondition1ComboBox.Items.Add("(none)");
            ExtraCondition2ComboBox.Items.Add("(none)");

            foreach (var header in fileInfo.HeaderRow)
            {
                Condition1ComboBox.Items.Add(header);
                Condition2ComboBox.Items.Add(header);
                ExtraCondition1ComboBox.Items.Add(header);
                ExtraCondition2ComboBox.Items.Add(header);
                ResultColumnComboBox.Items.Add(header);
                NgColumnComboBox.Items.Add(header);
                InputColumnComboBox.Items.Add(header);
            }

            var columnCount = fileInfo.HeaderRow.Count;
            Condition1ComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedCondition1Index, columnCount)
                ? fileInfo.SavedCondition1Index
                : 0;
            Condition2ComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedCondition2Index, columnCount)
                ? fileInfo.SavedCondition2Index
                : Math.Min(1, columnCount - 1);
            ExtraCondition1ComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedExtraCondition1Index, columnCount)
                ? fileInfo.SavedExtraCondition1Index + 1
                : 0;
            ExtraCondition2ComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedExtraCondition2Index, columnCount)
                ? fileInfo.SavedExtraCondition2Index + 1
                : 0;
            ResultColumnComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedResultIndex, columnCount)
                ? fileInfo.SavedResultIndex
                : (columnCount > 2 ? 2 : columnCount - 1);
            NgColumnComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedNgColumnIndex, columnCount)
                ? fileInfo.SavedNgColumnIndex
                : 0;
            InputColumnComboBox.SelectedIndex = IsValidIndex(fileInfo.SavedInputColumnIndex, columnCount)
                ? fileInfo.SavedInputColumnIndex
                : Math.Min(1, columnCount - 1);

            if (fileInfo.UseFormulaResult)
            {
                FormulaResultRadio.IsChecked = true;
            }
            else
            {
                DirectResultRadio.IsChecked = true;
            }

            ApplyResultModeUiState(fileInfo.UseFormulaResult);

            fileInfo.SavedCondition1Index = Condition1ComboBox.SelectedIndex;
            fileInfo.SavedCondition2Index = Condition2ComboBox.SelectedIndex;
            fileInfo.SavedExtraCondition1Index = ExtraCondition1ComboBox.SelectedIndex - 1;
            fileInfo.SavedExtraCondition2Index = ExtraCondition2ComboBox.SelectedIndex - 1;
            fileInfo.SavedResultIndex = ResultColumnComboBox.SelectedIndex;
            fileInfo.SavedNgColumnIndex = NgColumnComboBox.SelectedIndex;
            fileInfo.SavedInputColumnIndex = InputColumnComboBox.SelectedIndex;
        }

        public void HandleWebDroppedFiles(IReadOnlyList<string> filePaths)
        {
            string? firstTxtFile = filePaths.FirstOrDefault(f =>
                string.Equals(Path.GetExtension(f), ".txt", StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(firstTxtFile))
            {
                LoadFile(firstTxtFile);
            }
        }

        private static bool IsValidIndex(int index, int itemCount)
        {
            return index >= 0 && index < itemCount;
        }

        private void ResultModeChanged(object sender, RoutedEventArgs e)
        {
            var useFormula = sender == FormulaResultRadio || FormulaResultRadio?.IsChecked == true;
            ApplyResultModeUiState(useFormula);

            if (_currentFile != null)
            {
                _currentFile.UseFormulaResult = useFormula;
            }
        }

        private void ApplyResultModeUiState(bool useFormula)
        {
            if (ResultColumnComboBox == null || NgColumnComboBox == null || InputColumnComboBox == null)
            {
                return;
            }

            ResultColumnComboBox.IsEnabled = !useFormula;
            NgColumnComboBox.IsEnabled = useFormula;
            InputColumnComboBox.IsEnabled = useFormula;
        }

        private void DelimiterChanged(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                return;
            }

            if (TabDelimiterRadio.IsChecked == true)
            {
                _currentFile.Delimiter = "\t";
            }
            else if (CommaDelimiterRadio.IsChecked == true)
            {
                _currentFile.Delimiter = ",";
            }
            else if (SpaceDelimiterRadio.IsChecked == true)
            {
                _currentFile.Delimiter = " ";
            }

            LoadFileData(_currentFile);
        }

        private void HeaderRowChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentFile == null)
            {
                return;
            }

            if (!int.TryParse(HeaderRowTextBox.Text, out int headerRow) || headerRow <= 0)
            {
                return;
            }

            _currentFile.HeaderRowNumber = headerRow;
            LoadFileData(_currentFile);
        }

        private void ColumnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_currentFile == null)
            {
                return;
            }

            _currentFile.SavedCondition1Index = Condition1ComboBox.SelectedIndex;
            _currentFile.SavedCondition2Index = Condition2ComboBox.SelectedIndex;
            _currentFile.SavedExtraCondition1Index = ExtraCondition1ComboBox.SelectedIndex - 1;
            _currentFile.SavedExtraCondition2Index = ExtraCondition2ComboBox.SelectedIndex - 1;
            _currentFile.SavedResultIndex = ResultColumnComboBox.SelectedIndex;
            _currentFile.SavedNgColumnIndex = NgColumnComboBox.SelectedIndex;
            _currentFile.SavedInputColumnIndex = InputColumnComboBox.SelectedIndex;
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            _currentFile = null;

            PreviewGraphViewBase.ResetPreviewSummary(CurrentFileNameText, RowCountText, ColumnCountText, StatusText, "(No file)", "Load a file and select columns.");
            DataPreviewGrid.ItemsSource = null;

            Condition1ComboBox.Items.Clear();
            Condition2ComboBox.Items.Clear();
            ExtraCondition1ComboBox.Items.Clear();
            ExtraCondition2ComboBox.Items.Clear();
            ResultColumnComboBox.Items.Clear();
            NgColumnComboBox.Items.Clear();
            InputColumnComboBox.Items.Clear();
            DirectResultRadio.IsChecked = true;
            ApplyResultModeUiState(false);

            HeatMapModel = GraphSampleModelFactory.CreateHeatMapSample();
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Load data first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var state = new HeatMapReportState
            {
                Name = _currentFile.Name,
                FilePath = _currentFile.FilePath,
                Delimiter = _currentFile.Delimiter,
                HeaderRowNumber = _currentFile.HeaderRowNumber,
                RawData = GraphReportStorageHelper.CaptureRawTableData(_currentFile.FullData),
                SavedCondition1Index = _currentFile.SavedCondition1Index,
                SavedCondition2Index = _currentFile.SavedCondition2Index,
                SavedExtraCondition1Index = _currentFile.SavedExtraCondition1Index,
                SavedExtraCondition2Index = _currentFile.SavedExtraCondition2Index,
                SavedResultIndex = _currentFile.SavedResultIndex,
                SavedNgColumnIndex = _currentFile.SavedNgColumnIndex,
                SavedInputColumnIndex = _currentFile.SavedInputColumnIndex,
                UseFormulaResult = _currentFile.UseFormulaResult
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "heatmap.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            HeatMapReportState? state = GraphReportFileDialogHelper.LoadState<HeatMapReportState>("Load Graph Report");
            if (state == null)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _currentFile = new HeatMapFileInfo
            {
                Name = string.IsNullOrWhiteSpace(state.Name) ? "Loaded HeatMap Report" : state.Name,
                FilePath = state.FilePath,
                Delimiter = string.IsNullOrWhiteSpace(state.Delimiter) ? "\t" : state.Delimiter,
                HeaderRowNumber = Math.Max(1, state.HeaderRowNumber),
                FullData = GraphReportStorageHelper.BuildTableFromRawData(state.RawData),
                SavedCondition1Index = state.SavedCondition1Index,
                SavedCondition2Index = state.SavedCondition2Index,
                SavedExtraCondition1Index = state.SavedExtraCondition1Index,
                SavedExtraCondition2Index = state.SavedExtraCondition2Index,
                SavedResultIndex = state.SavedResultIndex,
                SavedNgColumnIndex = state.SavedNgColumnIndex,
                SavedInputColumnIndex = state.SavedInputColumnIndex,
                UseFormulaResult = state.UseFormulaResult
            };

            if (_currentFile.Delimiter == "\t")
            {
                TabDelimiterRadio.IsChecked = true;
            }
            else if (_currentFile.Delimiter == ",")
            {
                CommaDelimiterRadio.IsChecked = true;
            }
            else
            {
                SpaceDelimiterRadio.IsChecked = true;
            }

            HeaderRowTextBox.Text = _currentFile.HeaderRowNumber.ToString(CultureInfo.InvariantCulture);
            LoadFileData(_currentFile);
            HeatMapModel = GraphSampleModelFactory.CreateHeatMapSample();
        }

        private void GenerateHeatMapButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null || _currentFile.HeaderRow == null)
            {
                MessageBox.Show("Please load a TXT file first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var condition1Index = _currentFile.SavedCondition1Index;
            var condition2Index = _currentFile.SavedCondition2Index;
            var extraCondition1Index = _currentFile.SavedExtraCondition1Index;
            var extraCondition2Index = _currentFile.SavedExtraCondition2Index;
            var useFormulaResult = _currentFile.UseFormulaResult;
            var resultIndex = _currentFile.SavedResultIndex;
            var ngColumnIndex = _currentFile.SavedNgColumnIndex;
            var inputColumnIndex = _currentFile.SavedInputColumnIndex;

            if (condition1Index < 0 || condition2Index < 0)
            {
                MessageBox.Show("Select both Condition1 and Condition2.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!useFormulaResult && resultIndex < 0)
            {
                MessageBox.Show("In direct result mode, select a result column.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (useFormulaResult && (ngColumnIndex < 0 || inputColumnIndex < 0))
            {
                MessageBox.Show("In formula mode, select both NG and INPUT columns.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var duplicate = useFormulaResult
                ? HasDuplicateSelectedColumns(
                    condition1Index,
                    condition2Index,
                    extraCondition1Index,
                    extraCondition2Index,
                    ngColumnIndex,
                    inputColumnIndex)
                : HasDuplicateSelectedColumns(
                    condition1Index,
                    condition2Index,
                    extraCondition1Index,
                    extraCondition2Index,
                    resultIndex);

            if (duplicate)
            {
                MessageBox.Show(" /  Duplicate  .", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!TryBuildHeatMapData(
                    _currentFile.FullData,
                    condition1Index,
                    condition2Index,
                    extraCondition1Index,
                    extraCondition2Index,
                    useFormulaResult,
                    resultIndex,
                    ngColumnIndex,
                    inputColumnIndex,
                    out var matrix,
                    out var xLabels,
                    out var yLabels,
                    out var annotationTexts,
                    out var usedRows,
                    out var skippedRows,
                    out var minimumValue,
                    out var maximumValue,
                    out var errorMessage))
            {
                MessageBox.Show(errorMessage, "Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var condition1Name = _currentFile.HeaderRow[condition1Index];
            var condition2Name = _currentFile.HeaderRow[condition2Index];
            var extraCondition1Name = extraCondition1Index >= 0 ? _currentFile.HeaderRow[extraCondition1Index] : null;
            var extraCondition2Name = extraCondition2Index >= 0 ? _currentFile.HeaderRow[extraCondition2Index] : null;
            var resultName = useFormulaResult
                ? $"{_currentFile.HeaderRow[ngColumnIndex]} * 1000000 / {_currentFile.HeaderRow[inputColumnIndex]}"
                : _currentFile.HeaderRow[resultIndex];
            var xAxisTitle = BuildAxisTitle(condition1Name, extraCondition1Name);
            var yAxisTitle = BuildAxisTitle(condition2Name, extraCondition2Name);

            HeatMapModel = CreateHeatMapPlot(
                matrix,
                xLabels,
                yLabels,
                xAxisTitle,
                yAxisTitle,
                resultName,
                minimumValue,
                maximumValue,
                annotationTexts);

            var popupModel = CreateHeatMapPlot(
                matrix,
                xLabels,
                yLabels,
                xAxisTitle,
                yAxisTitle,
                resultName,
                minimumValue,
                maximumValue,
                annotationTexts);

            var viewerWindow = new HeatMapViewerWindow(popupModel);
            var ownerWindow = Window.GetWindow(this);
            if (ownerWindow is not null && ownerWindow.IsVisible)
            {
                viewerWindow.Owner = ownerWindow;
            }

            viewerWindow.Show();

            var modeText = useFormulaResult ? "Formula mode" : "Direct result mode";
            StatusText.Text = $"Heatmap  (  , {modeText}), valid rows {usedRows:N0}, numeric-conversion-failed/empty skipped {skippedRows:N0}.";
        }

        private void OpenFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                MessageBox.Show("Load a file first.", "Notice",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            LoadFileData(_currentFile);
        }

        private static bool HasDuplicateSelectedColumns(params int[] selectedColumns)
        {
            var distinct = selectedColumns.Where(x => x >= 0).Distinct().Count();
            var total = selectedColumns.Count(x => x >= 0);
            return distinct != total;
        }

        private static string BuildAxisTitle(string mainName, string? extraName)
        {
            return string.IsNullOrWhiteSpace(extraName)
                ? mainName
                : $"{mainName}+{extraName}";
        }

        private static PlotModel CreateHeatMapPlot(
            double[,] matrix,
            IReadOnlyList<string> xLabels,
            IReadOnlyList<string> yLabels,
            string xAxisTitle,
            string yAxisTitle,
            string resultName,
            double minimumValue,
            double maximumValue,
            string?[,] annotationTexts)
        {
            var model = new PlotModel
            {
                Title = $"{resultName} Heatmap ({xAxisTitle} x {yAxisTitle})",
                Padding = new OxyThickness(10, 8, 10, 8)
            };
            var palette = OxyPalettes.BlueWhiteRed(200);

            var xAngle = xLabels.Count > 8 || xLabels.Any(label => label.Length > 8) ? -35 : 0;

            var xAxis = new CategoryAxis
            {
                Position = AxisPosition.Bottom,
                Title = xAxisTitle,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                Angle = xAngle,
                AxisTitleDistance = 12,
                GapWidth = 0,
                TickStyle = OxyPlot.Axes.TickStyle.None,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.None,
                MajorGridlineColor = OxyColor.FromAColor(80, OxyColors.LightGray),
                AxislineStyle = LineStyle.Solid
            };
            foreach (var label in xLabels)
            {
                xAxis.Labels.Add(label);
            }
            model.Axes.Add(xAxis);

            var yAxis = new CategoryAxis
            {
                Position = AxisPosition.Left,
                Title = yAxisTitle,
                IsPanEnabled = false,
                IsZoomEnabled = false,
                AxisTitleDistance = 12,
                GapWidth = 0,
                TickStyle = OxyPlot.Axes.TickStyle.None,
                MajorGridlineStyle = LineStyle.Solid,
                MinorGridlineStyle = LineStyle.None,
                MajorGridlineColor = OxyColor.FromAColor(80, OxyColors.LightGray),
                AxislineStyle = LineStyle.Solid
            };
            foreach (var label in yLabels)
            {
                yAxis.Labels.Add(label);
            }
            model.Axes.Add(yAxis);

            model.Axes.Add(new LinearColorAxis
            {
                Position = AxisPosition.Right,
                Title = resultName,
                Palette = palette,
                Minimum = minimumValue,
                Maximum = maximumValue,
                AxisDistance = 8,
                InvalidNumberColor = OxyColors.White
            });

            const double x0 = -0.5;
            var x1 = xLabels.Count - 0.5;
            const double y0 = -0.5;
            var y1 = yLabels.Count - 0.5;

            model.Series.Add(new HeatMapSeries
            {
                X0 = x0,
                X1 = x1,
                Y0 = y0,
                Y1 = y1,
                Data = matrix,
                CoordinateDefinition = HeatMapCoordinateDefinition.Edge,
                Interpolate = false,
                RenderMethod = HeatMapRenderMethod.Rectangles
            });

            AddHeatMapValueAnnotations(model, matrix, palette, minimumValue, maximumValue, x0, x1, y0, y1, annotationTexts);
            return model;
        }

        private static void AddHeatMapValueAnnotations(
            PlotModel model,
            double[,] matrix,
            OxyPalette palette,
            double minimumValue,
            double maximumValue,
            double x0,
            double x1,
            double y0,
            double y1,
            string?[,]? annotationTexts)
        {
            var xCount = matrix.GetLength(0);
            var yCount = matrix.GetLength(1);
            if (xCount == 0 || yCount == 0)
            {
                return;
            }

            var cellWidth = (x1 - x0) / xCount;
            var cellHeight = (y1 - y0) / yCount;
            var cellCount = xCount * yCount;

            var fontSize = cellCount switch
            {
                <= 80 => 11,
                <= 180 => 10,
                <= 320 => 9,
                _ => 8
            };

            for (var x = 0; x < xCount; x++)
            {
                for (var y = 0; y < yCount; y++)
                {
                    var value = matrix[x, y];
                    if (double.IsNaN(value))
                    {
                        continue;
                    }

                    model.Annotations.Add(new TextAnnotation
                    {
                        Text = BuildHeatMapAnnotationText(annotationTexts, x, y, value),
                        TextPosition = new DataPoint(
                            x0 + ((x + 0.5) * cellWidth),
                            y0 + ((y + 0.5) * cellHeight)),
                        TextHorizontalAlignment = OxyPlot.HorizontalAlignment.Center,
                        TextVerticalAlignment = OxyPlot.VerticalAlignment.Middle,
                        TextColor = GetHeatMapTextColor(value, minimumValue, maximumValue, palette),
                        FontSize = fontSize,
                        FontWeight = OxyPlot.FontWeights.Bold,
                        Offset = new ScreenVector(0, 0),
                        Padding = new OxyThickness(0),
                        Stroke = OxyColors.Transparent,
                        StrokeThickness = 0,
                        Layer = AnnotationLayer.AboveSeries,
                        ClipByXAxis = true,
                        ClipByYAxis = true
                    });
                }
            }
        }

        private static string BuildHeatMapAnnotationText(string?[,]? annotationTexts, int x, int y, double value)
        {
            if (annotationTexts != null &&
                x >= 0 && x < annotationTexts.GetLength(0) &&
                y >= 0 && y < annotationTexts.GetLength(1) &&
                !string.IsNullOrWhiteSpace(annotationTexts[x, y]))
            {
                return annotationTexts[x, y]!;
            }

            return FormatHeatMapCellValue(value);
        }

        private static string FormatHeatMapCellValue(double value)
        {
            var truncated = Math.Truncate(value);
            return truncated.ToString("N0", CultureInfo.InvariantCulture);
        }

        private static OxyColor GetHeatMapTextColor(
            double value,
            double minimumValue,
            double maximumValue,
            OxyPalette palette)
        {
            if (palette.Colors.Count == 0)
            {
                return OxyColors.Black;
            }

            var range = maximumValue - minimumValue;
            if (Math.Abs(range) < 1e-12)
            {
                return OxyColors.Black;
            }

            var ratio = (value - minimumValue) / range;
            ratio = Math.Max(0, Math.Min(1, ratio));
            var paletteIndex = (int)Math.Round(ratio * (palette.Colors.Count - 1));
            var baseColor = palette.Colors[paletteIndex];

            var brightness = (0.299 * baseColor.R) + (0.587 * baseColor.G) + (0.114 * baseColor.B);
            return brightness < 135 ? OxyColors.White : OxyColors.Black;
        }

        private static string FormatCategoryLabel(double value, IReadOnlyList<string> labels)
        {
            var rounded = (int)Math.Round(value);
            if (rounded < 0 || rounded >= labels.Count)
            {
                return string.Empty;
            }

            if (Math.Abs(value - rounded) > 0.25)
            {
                return string.Empty;
            }

            var label = labels[rounded];
            return label.Length <= 20 ? label : label[..20] + "...";
        }

        private static bool TryBuildHeatMapData(
            DataTable dataTable,
            int condition1Index,
            int condition2Index,
            int extraCondition1Index,
            int extraCondition2Index,
            bool useFormulaResult,
            int resultIndex,
            int ngColumnIndex,
            int inputColumnIndex,
            out double[,] matrix,
            out List<string> xLabels,
            out List<string> yLabels,
            out string?[,] annotationTexts,
            out int usedRows,
            out int skippedRows,
            out double minimumValue,
            out double maximumValue,
            out string errorMessage)
        {
            matrix = new double[0, 0];
            xLabels = new List<string>();
            yLabels = new List<string>();
            annotationTexts = new string?[0, 0];
            usedRows = 0;
            skippedRows = 0;
            minimumValue = 0;
            maximumValue = 0;
            errorMessage = string.Empty;

            var grouped = new Dictionary<(string X, string Y), List<double>>();
            var groupedFormulaSums = new Dictionary<(string X, string Y), (double NgSum, double InputSum)>();
            var xValues = new HashSet<string>(StringComparer.Ordinal);
            var yValues = new HashSet<string>(StringComparer.Ordinal);

            foreach (DataRow row in dataTable.Rows)
            {
                if (!TryBuildConditionValue(row, condition1Index, extraCondition1Index, out var condition1Value) ||
                    !TryBuildConditionValue(row, condition2Index, extraCondition2Index, out var condition2Value))
                {
                    continue;
                }

                if (!TryResolveResultValue(
                        row,
                        useFormulaResult,
                        resultIndex,
                        ngColumnIndex,
                        inputColumnIndex,
                        out var numericResult,
                        out var ngValue,
                        out var inputValue))
                {
                    skippedRows++;
                    continue;
                }

                usedRows++;
                xValues.Add(condition1Value);
                yValues.Add(condition2Value);

                var key = (condition1Value, condition2Value);
                if (!grouped.TryGetValue(key, out var values))
                {
                    values = new List<double>();
                    grouped[key] = values;
                }

                values.Add(numericResult);

                if (useFormulaResult)
                {
                    if (!groupedFormulaSums.TryGetValue(key, out var sums))
                    {
                        sums = (0, 0);
                    }

                    groupedFormulaSums[key] = (sums.NgSum + ngValue, sums.InputSum + inputValue);
                }
            }

            if (grouped.Count == 0)
            {
                errorMessage = "No numeric data found in selected result column.";
                return false;
            }

            xLabels = OrderLabels(xValues);
            yLabels = OrderLabels(yValues);

            var xIndexMap = xLabels.Select((label, index) => new { label, index })
                .ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);
            var yIndexMap = yLabels.Select((label, index) => new { label, index })
                .ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);

            matrix = new double[xLabels.Count, yLabels.Count];
            annotationTexts = new string?[xLabels.Count, yLabels.Count];
            for (var x = 0; x < xLabels.Count; x++)
            {
                for (var y = 0; y < yLabels.Count; y++)
                {
                    matrix[x, y] = double.NaN;
                }
            }

            var aggregatedValues = new List<double>(grouped.Count);
            foreach (var entry in grouped)
            {
                var xIndex = xIndexMap[entry.Key.X];
                var yIndex = yIndexMap[entry.Key.Y];
                var averageValue = entry.Value.Average();

                matrix[xIndex, yIndex] = averageValue;
                if (useFormulaResult && groupedFormulaSums.TryGetValue(entry.Key, out var formulaSums))
                {
                    annotationTexts[xIndex, yIndex] =
                        $"{FormatHeatMapCellValue(averageValue)}\n({FormatHeatMapCellValue(formulaSums.NgSum)}/{FormatHeatMapCellValue(formulaSums.InputSum)})";
                }
                else
                {
                    annotationTexts[xIndex, yIndex] = FormatHeatMapCellValue(averageValue);
                }

                aggregatedValues.Add(averageValue);
            }

            minimumValue = aggregatedValues.Min();
            maximumValue = aggregatedValues.Max();

            if (Math.Abs(maximumValue - minimumValue) < 1e-9)
            {
                minimumValue -= 0.5;
                maximumValue += 0.5;
            }

            return true;
        }

        private static bool TryBuildConditionValue(
            DataRow row,
            int mainIndex,
            int extraIndex,
            out string value)
        {
            value = string.Empty;

            var main = row[mainIndex]?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(main))
            {
                return false;
            }

            if (extraIndex < 0)
            {
                value = main;
                return true;
            }

            var extra = row[extraIndex]?.ToString()?.Trim();
            if (string.IsNullOrWhiteSpace(extra))
            {
                value = main;
                return true;
            }

            value = $"{main}|{extra}";
            return true;
        }

        private static bool TryResolveResultValue(
            DataRow row,
            bool useFormulaResult,
            int resultIndex,
            int ngColumnIndex,
            int inputColumnIndex,
            out double resultValue,
            out double ngValue,
            out double inputValue)
        {
            resultValue = 0;
            ngValue = 0;
            inputValue = 0;

            if (!useFormulaResult)
            {
                var raw = row[resultIndex]?.ToString();
                return GraphMakerParsingHelper.TryParseDouble(raw, out resultValue);
            }

            var ngRaw = row[ngColumnIndex]?.ToString();
            var inputRaw = row[inputColumnIndex]?.ToString();

            if (!GraphMakerParsingHelper.TryParseDouble(ngRaw, out ngValue))
            {
                return false;
            }

            if (!GraphMakerParsingHelper.TryParseDouble(inputRaw, out inputValue))
            {
                return false;
            }

            if (Math.Abs(inputValue) < 1e-12)
            {
                return false;
            }

            resultValue = ngValue * 1000000.0 / inputValue;
            return true;
        }

        private static List<string> OrderLabels(IEnumerable<string> labels)
        {
            var normalized = labels
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (normalized.Count == 0)
            {
                return new List<string>();
            }

            var allNumeric = normalized.All(label => GraphMakerParsingHelper.TryParseDouble(label, out _));
            if (allNumeric)
            {
                return normalized
                    .OrderBy(label => ParseDouble(label))
                    .ThenBy(label => label, StringComparer.Ordinal)
                    .ToList();
            }

            return normalized
                .OrderBy(label => label, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static double ParseDouble(string text)
        {
            return GraphMakerParsingHelper.TryParseDouble(text, out var value) ? value : double.NaN;
        }

        public object GetWebModuleSnapshot()
        {
            string delimiter = TabDelimiterRadio.IsChecked == true ? "tab"
                : CommaDelimiterRadio.IsChecked == true ? "comma"
                : "space";

            return new
            {
                moduleType = "GraphMakerHeatMap",
                fileName = CurrentFileNameText.Text ?? "(No file)",
                rowCount = RowCountText.Text ?? "0",
                columnCount = ColumnCountText.Text ?? "0",
                status = StatusText.Text ?? string.Empty,
                delimiter,
                headerRow = HeaderRowTextBox.Text ?? "1",
                useFormulaResult = FormulaResultRadio.IsChecked == true,
                resultMode = FormulaResultRadio.IsChecked == true ? "NG * 1000000 / INPUT" : "Use result column",
                condition1 = Condition1ComboBox.SelectedItem?.ToString() ?? string.Empty,
                condition2 = Condition2ComboBox.SelectedItem?.ToString() ?? string.Empty,
                extraCondition1 = ExtraCondition1ComboBox.SelectedItem?.ToString() ?? string.Empty,
                extraCondition2 = ExtraCondition2ComboBox.SelectedItem?.ToString() ?? string.Empty,
                resultColumn = ResultColumnComboBox.SelectedItem?.ToString() ?? string.Empty,
                ngColumn = NgColumnComboBox.SelectedItem?.ToString() ?? string.Empty,
                inputColumn = InputColumnComboBox.SelectedItem?.ToString() ?? string.Empty,
                options = Condition1ComboBox.Items.Cast<object>().Select(item => item?.ToString() ?? string.Empty).ToArray(),
                previewColumns = BuildPreviewColumns(_currentFile?.FullData),
                previewRows = BuildPreviewRows(_currentFile?.FullData)
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("delimiter", out JsonElement delimiterElement))
            {
                string delimiter = delimiterElement.GetString() ?? "tab";
                TabDelimiterRadio.IsChecked = delimiter == "tab";
                CommaDelimiterRadio.IsChecked = delimiter == "comma";
                SpaceDelimiterRadio.IsChecked = delimiter == "space";
                if (_currentFile != null)
                {
                    DelimiterChanged(this, new RoutedEventArgs());
                }
            }

            if (payload.TryGetProperty("headerRow", out JsonElement headerRowElement))
            {
                HeaderRowTextBox.Text = headerRowElement.GetString() ?? "1";
            }

            if (payload.TryGetProperty("useFormulaResult", out JsonElement formulaElement))
            {
                bool useFormula = formulaElement.GetBoolean();
                FormulaResultRadio.IsChecked = useFormula;
                DirectResultRadio.IsChecked = !useFormula;
            }

            SetSelection(payload, "condition1", Condition1ComboBox);
            SetSelection(payload, "condition2", Condition2ComboBox);
            SetSelection(payload, "extraCondition1", ExtraCondition1ComboBox);
            SetSelection(payload, "extraCondition2", ExtraCondition2ComboBox);
            SetSelection(payload, "resultColumn", ResultColumnComboBox);
            SetSelection(payload, "ngColumn", NgColumnComboBox);
            SetSelection(payload, "inputColumn", InputColumnComboBox);

            NotifyWebModuleSnapshotChanged();
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

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "browse-file":
                    BrowseButton_Click(this, new RoutedEventArgs());
                    break;
                case "remove-file":
                    RemoveFileButton_Click(this, new RoutedEventArgs());
                    break;
                case "save-report":
                    SaveReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "load-report":
                    LoadReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "generate-heatmap":
                    GenerateHeatMapButton_Click(this, new RoutedEventArgs());
                    break;
            }

            return GetWebModuleSnapshot();
        }

        private static void SetSelection(JsonElement payload, string propertyName, ComboBox comboBox)
        {
            if (!payload.TryGetProperty(propertyName, out JsonElement property))
            {
                return;
            }

            string value = property.GetString() ?? string.Empty;
            object? match = comboBox.Items.Cast<object>().FirstOrDefault(item => string.Equals(item?.ToString(), value, StringComparison.Ordinal));
            if (match is not null)
            {
                comboBox.SelectedItem = match;
            }
        }

        private void NotifyWebModuleSnapshotChanged()
        {
            WebModuleSnapshotChanged?.Invoke();
        }
    }
}


