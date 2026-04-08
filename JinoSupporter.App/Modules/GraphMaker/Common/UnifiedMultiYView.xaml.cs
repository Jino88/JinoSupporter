using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Win32;
using OxyPlot;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GraphMaker
{
    public partial class UnifiedMultiYView : GraphViewBase
    {
        private sealed class UnifiedMultiYReportState
        {
            public List<UnifiedMultiYFileState> Files { get; set; } = new();
            public string? SelectedFilePath { get; set; }
            public string SpecText { get; set; } = string.Empty;
            public string UpperText { get; set; } = string.Empty;
            public string LowerText { get; set; } = string.Empty;
            public bool ApplyToAllColumns { get; set; } = true;
        }

        private sealed class UnifiedMultiYFileState
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public string Delimiter { get; set; } = "\t";
            public int HeaderRowNumber { get; set; } = 1;
            public MultiYInputMode Mode { get; set; }
            public RawTableData RawData { get; set; } = new();
            public Dictionary<string, ColumnLimitSetting> SavedColumnLimits { get; set; } = new(StringComparer.OrdinalIgnoreCase);
            public List<string> SelectedColumns { get; set; } = new();
        }

        private readonly ObservableCollection<FileInfo_DailySampling> _loadedFiles = new();
        private readonly ObservableCollection<SelectableColumnOption> _columnOptions = new();
        private readonly Dictionary<string, HashSet<string>> _selectedColumnsByFile = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, MultiYInputMode> _inputModesByFile = new(StringComparer.OrdinalIgnoreCase);

        private FileInfo_DailySampling? _currentFile;
        private bool _combinedYAxisView;
        private bool _isApplyingAutoDetect;
        private bool _isApplyingModeUi;
        private string _combinedXAxisColumn = string.Empty;
        private bool _combinedXAxisIsDate = true;
        public UnifiedMultiYView()
        {
            InitializeComponent();
            FileListBox.ItemsSource = _loadedFiles;
            ColumnOptionListBox.ItemsSource = _columnOptions;
            NotifyWebModuleSnapshotChanged();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Select Data Files",
                Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            foreach (string filePath in dialog.FileNames)
            {
                LoadFileDirect(filePath);
            }
        }

        private void LoadFileDirect(string filePath)
        {
            if (_loadedFiles.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
            {
                FileListBox.SelectedItem = _loadedFiles.First(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
                return;
            }

            string delimiter = GraphMakerTableHelper.DetectDelimiter(filePath);
            _isApplyingAutoDetect = true;
            TabDelimiterRadio.IsChecked = delimiter == "\t";
            CommaDelimiterRadio.IsChecked = delimiter == ",";
            SpaceDelimiterRadio.IsChecked = delimiter == " ";
            if (TabDelimiterRadio.IsChecked != true && CommaDelimiterRadio.IsChecked != true && SpaceDelimiterRadio.IsChecked != true)
            {
                TabDelimiterRadio.IsChecked = true;
                delimiter = "\t";
            }
            _isApplyingAutoDetect = false;

            MultiYInputMode mode = GetModeFromRadio();
            int headerRow = int.TryParse(HeaderRowTextBox?.Text, out int n) && n >= 1 ? n : 1;
            LoadFile(filePath, mode, delimiter, headerRow);
        }

        private MultiYInputMode GetModeFromRadio()
        {
            if (DateSingleYRadio?.IsChecked == true) return MultiYInputMode.DateSingleY;
            if (DateNoHeaderMultiYRadio?.IsChecked == true) return MultiYInputMode.DateNoHeaderMultiY;
            return MultiYInputMode.HeaderMultiY;
        }

        private void LoadFile(string filePath, MultiYInputMode mode, string delimiter, int headerRowNumber)
        {
            if (_loadedFiles.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            var file = GraphMakerFileViewHelper.CreateFileInfo(
                filePath,
                delimiter,
                mode == MultiYInputMode.DateNoHeaderMultiY ? 0 : Math.Max(1, headerRowNumber));

            _loadedFiles.Add(file);
            _inputModesByFile[file.FilePath] = mode;

            if (_loadedFiles.Count == 1)
            {
                FileListBox.SelectedItem = file;
            }
        }

        private void FileListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SaveCurrentSelectionState();

            _currentFile = FileListBox.SelectedItem as FileInfo_DailySampling;
            if (_currentFile == null)
            {
                DataPreviewGrid.ItemsSource = null;
                _columnOptions.Clear();
                UpdateModeUi(null);
                return;
            }

            UpdateModeUi(GetCurrentMode());
            LoadAndBindCurrentFile();
        }

        private MultiYInputMode? GetCurrentMode()
        {
            if (_currentFile == null)
            {
                return null;
            }

            return _inputModesByFile.TryGetValue(_currentFile.FilePath, out MultiYInputMode mode)
                ? mode
                : MultiYInputMode.HeaderMultiY;
        }

        private void UpdateModeUi(MultiYInputMode? mode)
        {
            _isApplyingModeUi = true;
            if (!mode.HasValue)
            {
                HeaderMultiYRadio.IsChecked = true;
                ModeHintTextBlock.Text = "Load a file and choose how it should be interpreted.";
                HeaderRowPanel.Visibility = Visibility.Visible;
                _isApplyingModeUi = false;
                return;
            }

            HeaderMultiYRadio.IsChecked = mode.Value == MultiYInputMode.HeaderMultiY;
            DateSingleYRadio.IsChecked = mode.Value == MultiYInputMode.DateSingleY;
            DateNoHeaderMultiYRadio.IsChecked = mode.Value == MultiYInputMode.DateNoHeaderMultiY;
            ModeHintTextBlock.Text = mode.Value switch
            {
                MultiYInputMode.HeaderMultiY => "Header row can be changed. All columns are treated as Y columns.",
                MultiYInputMode.DateSingleY => "First column must be Date. Only one Y column should be selected.",
                _ => "No header row. First column is Date; remaining become Value1, Value2, ..."
            };
            HeaderRowPanel.Visibility = mode.Value == MultiYInputMode.DateNoHeaderMultiY ? Visibility.Collapsed : Visibility.Visible;
            _isApplyingModeUi = false;
        }

        private void ModeRadioChanged(object sender, RoutedEventArgs e)
        {
            if (_isApplyingModeUi || _currentFile == null || !IsLoaded)
            {
                return;
            }

            MultiYInputMode newMode = GetModeFromRadio();
            _inputModesByFile[_currentFile.FilePath] = newMode;
            _currentFile.HeaderRowNumber = newMode == MultiYInputMode.DateNoHeaderMultiY
                ? 0
                : (int.TryParse(HeaderRowTextBox?.Text, out int n) && n >= 1 ? n : 1);
            HeaderRowPanel.Visibility = newMode == MultiYInputMode.DateNoHeaderMultiY ? Visibility.Collapsed : Visibility.Visible;
            ModeHintTextBlock.Text = newMode switch
            {
                MultiYInputMode.HeaderMultiY => "Header row can be changed. All columns are treated as Y columns.",
                MultiYInputMode.DateSingleY => "First column must be Date. Only one Y column should be selected.",
                _ => "No header row. First column is Date; remaining become Value1, Value2, ..."
            };
            _currentFile.FullData = null!;
            LoadAndBindCurrentFile();
        }

        private void LoadAndBindCurrentFile()
        {
            if (_currentFile == null)
            {
                return;
            }

            try
            {
                if (GetCurrentMode() == MultiYInputMode.DateNoHeaderMultiY)
                {
                    GraphMakerFileViewHelper.ParseDateNoHeaderIntoDataTable(_currentFile);
                    GraphMakerFileViewHelper.EnsureColumnLimits(
                        _currentFile,
                        name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase));
                }
                else
                {
                    GraphMakerFileViewHelper.ParseHeaderIntoDataTable(_currentFile, 1);
                    GraphMakerFileViewHelper.EnsureColumnLimits(_currentFile);
                }

                BindColumnOptions(_currentFile);
                BindPreviewGrid(_currentFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BindPreviewGrid(FileInfo_DailySampling file)
        {
            if (file.FullData == null)
            {
                DataPreviewGrid.ItemsSource = null;
                return;
            }

            if (GetCurrentMode() == MultiYInputMode.DateSingleY)
            {
                IReadOnlyCollection<string> selectedColumns = _selectedColumnsByFile.TryGetValue(file.FilePath, out HashSet<string>? selected)
                    ? selected.Where(name => file.FullData.Columns.Contains(name)).ToList()
                    : file.FullData.Columns.Cast<DataColumn>()
                        .Select(column => column.ColumnName)
                        .Where(name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                DataTable previewTable = BuildSingleYTable(file, selectedColumns);
                DataPreviewGrid.ItemsSource = previewTable.DefaultView;
                return;
            }

            DataPreviewGrid.ItemsSource = file.FullData.DefaultView;
        }

        private void BindColumnOptions(FileInfo_DailySampling file)
        {
            bool excludeDate = GetCurrentMode() != MultiYInputMode.HeaderMultiY;
            GraphMakerFileViewHelper.BindColumnOptions(
                file,
                _columnOptions,
                _selectedColumnsByFile,
                name => !excludeDate || !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrWhiteSpace(_combinedXAxisColumn) ||
                !_columnOptions.Any(option => string.Equals(option.ColumnName, _combinedXAxisColumn, StringComparison.OrdinalIgnoreCase)))
            {
                _combinedXAxisColumn = _columnOptions.FirstOrDefault()?.ColumnName ?? string.Empty;
            }
        }

        private void SaveCurrentSelectionState()
        {
            GraphMakerFileViewHelper.SaveCurrentSelectionState(_currentFile, _columnOptions, _selectedColumnsByFile);
            if (_currentFile != null)
            {
                BindPreviewGrid(_currentFile);
            }
        }

        private void ApplyLimitsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                return;
            }

            string spec = SpecValueTextBox.Text?.Trim() ?? string.Empty;
            string upper = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty;
            string lower = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty;

            IEnumerable<string> targetColumns = ApplyToAllColumnsCheckBox.IsChecked == true
                ? _columnOptions.Select(c => c.ColumnName)
                : _columnOptions.Where(option => option.IsSelected).Select(option => option.ColumnName);

            foreach (string columnName in targetColumns)
            {
                if (!_currentFile.SavedColumnLimits.TryGetValue(columnName, out ColumnLimitSetting? limit))
                {
                    limit = new ColumnLimitSetting { ColumnName = columnName };
                    _currentFile.SavedColumnLimits[columnName] = limit;
                }

                limit.SpecValue = spec;
                limit.UpperValue = upper;
                limit.LowerValue = lower;
            }
        }

        private void GenerateGraphButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                MessageBox.Show("Please load and select a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SaveCurrentSelectionState();
            var selectedColumns = _selectedColumnsByFile.TryGetValue(_currentFile.FilePath, out HashSet<string>? selectedSet)
                ? selectedSet.ToList()
                : _columnOptions.Select(c => c.ColumnName).ToList();

            if (selectedColumns.Count == 0)
            {
                MessageBox.Show("Select at least one data column to plot.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (_combinedYAxisView)
            {
                if (string.IsNullOrWhiteSpace(_combinedXAxisColumn))
                {
                    MessageBox.Show("Select one Y column to use as the combined X axis first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                List<string> combinedYColumns = selectedColumns
                    .Where(column => !string.Equals(column, _combinedXAxisColumn, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (combinedYColumns.Count == 0)
                {
                    MessageBox.Show("Select at least one Y column besides the combined X axis.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                FileInfo_DailySampling combinedFile = BuildCombinedFile(_currentFile, _combinedXAxisColumn, combinedYColumns, _combinedXAxisIsDate);
                var combinedResult = MultiColumnGraphCalculator.Calculate(combinedFile, new List<string> { "Value" });
                var combinedWindow = new MultiColumnResultWindow(_currentFile.Name, _combinedXAxisColumn, combinedResult)
                {
                    WindowState = WindowState.Maximized
                };
                Window? combinedOwnerWindow = Window.GetWindow(this);
                if (combinedOwnerWindow is not null && combinedOwnerWindow.IsVisible)
                {
                    combinedWindow.Owner = combinedOwnerWindow;
                    combinedWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                }
                combinedWindow.Show();
                return;
            }

            if (GetCurrentMode() == MultiYInputMode.DateNoHeaderMultiY)
            {
                FileInfo_DailySampling combinedFile = BuildCombinedFile(_currentFile, selectedColumns);
                var result = MultiColumnGraphCalculator.Calculate(combinedFile, new List<string> { "Value" });
                var multiColumnWindow = new MultiColumnResultWindow(_currentFile.Name, "Date", result)
                {
                    WindowState = WindowState.Maximized
                };
                Window? ownerWindow = Window.GetWindow(this);
                if (ownerWindow is not null && ownerWindow.IsVisible)
                {
                    multiColumnWindow.Owner = ownerWindow;
                    multiColumnWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                }
                multiColumnWindow.Show();
                return;
            }

            if (GetCurrentMode() == MultiYInputMode.DateSingleY)
            {
                if (!ValidateFirstColumnDate(_currentFile.FullData, out string? errorMessage))
                {
                    MessageBox.Show(errorMessage, "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                DataTable singleTable = BuildSingleYTable(_currentFile, selectedColumns);
                if (singleTable.Rows.Count == 0)
                {
                    MessageBox.Show("No valid Date/Value rows were found.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _currentFile.SavedColumnLimits["Value"] = new ColumnLimitSetting
                {
                    ColumnName = "Value",
                    SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty,
                    UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty,
                    LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty
                };

                var graphDataList = new List<DailySamplingGraphData>
                {
                    new DailySamplingGraphData
                    {
                        FileName = _currentFile.Name,
                        Dates = singleTable.Rows.Cast<DataRow>().Select(row => row["Date"]?.ToString() ?? string.Empty).ToList(),
                        SampleNumbers = new List<string> { "Value" },
                        Data = singleTable,
                        UpperLimitValue = _currentFile.SavedColumnLimits.TryGetValue("Value", out ColumnLimitSetting? limitValue) &&
                                          GraphMakerParsingHelper.TryParseDouble(limitValue.UpperValue, out double usl) ? usl : null,
                        SpecValue = _currentFile.SavedColumnLimits.TryGetValue("Value", out limitValue) &&
                                    GraphMakerParsingHelper.TryParseDouble(limitValue.SpecValue, out double spec) ? spec : null,
                        LowerLimitValue = _currentFile.SavedColumnLimits.TryGetValue("Value", out limitValue) &&
                                          GraphMakerParsingHelper.TryParseDouble(limitValue.LowerValue, out double lsl) ? lsl : null,
                        DataColor = Colors.Black,
                        SpecColor = Colors.Black,
                        UpperColor = Colors.Black,
                        LowerColor = Colors.Black,
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
                return;
            }

            var noXResult = NoXMultiYGraphCalculator.Calculate(_currentFile, selectedColumns);
            var noXWindow = new NoXMultiYResultWindow(_currentFile.Name, noXResult)
            {
                WindowState = WindowState.Maximized
            };
            Window? noXOwnerWindow = Window.GetWindow(this);
            if (noXOwnerWindow is not null && noXOwnerWindow.IsVisible)
            {
                noXWindow.Owner = noXOwnerWindow;
                noXWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            noXWindow.Show();
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (FileListBox.SelectedItem is not FileInfo_DailySampling selected)
            {
                return;
            }

            _loadedFiles.Remove(selected);
            _selectedColumnsByFile.Remove(selected.FilePath);
            _inputModesByFile.Remove(selected.FilePath);

            if (_loadedFiles.Count == 0)
            {
                _currentFile = null;
                DataPreviewGrid.ItemsSource = null;
                _columnOptions.Clear();
                UpdateModeUi(null);
                return;
            }

            FileListBox.SelectedIndex = 0;
        }

        private void DelimiterChanged(object sender, RoutedEventArgs e)
        {
            if (_isApplyingAutoDetect || _currentFile == null || !IsLoaded)
            {
                return;
            }

            _currentFile.Delimiter = TabDelimiterRadio.IsChecked == true ? "\t" : CommaDelimiterRadio.IsChecked == true ? "," : " ";
            _currentFile.FullData = null!;
            LoadAndBindCurrentFile();
        }

        private void HeaderRowChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentFile == null || GetCurrentMode() != MultiYInputMode.HeaderMultiY)
            {
                return;
            }

            if (!int.TryParse(HeaderRowTextBox.Text, out int headerRow) || headerRow <= 0)
            {
                return;
            }

            _currentFile.HeaderRowNumber = headerRow;
            _currentFile.FullData = null!;
            LoadAndBindCurrentFile();
        }

        private void DropZone_Drop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            if (e.Data.GetData(DataFormats.FileDrop) is not string[] files)
            {
                return;
            }

            foreach (string file in files)
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".txt" || ext == ".csv")
                {
                    LoadFileDirect(file);
                }
            }
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

            if (accepted.Length == 0)
            {
                return;
            }

            foreach (string file in accepted)
            {
                LoadFileDirect(file);
            }

            NotifyWebModuleSnapshotChanged();
        }

        private void SaveReportButton_Click(object sender, RoutedEventArgs e)
        {
            SaveCurrentSelectionState();
            if (_loadedFiles.Count == 0)
            {
                MessageBox.Show("Load at least one file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var state = new UnifiedMultiYReportState
            {
                SelectedFilePath = _currentFile?.FilePath,
                SpecText = SpecValueTextBox.Text ?? string.Empty,
                UpperText = UpperLimitValueTextBox.Text ?? string.Empty,
                LowerText = LowerLimitValueTextBox.Text ?? string.Empty,
                ApplyToAllColumns = ApplyToAllColumnsCheckBox.IsChecked == true,
                Files = _loadedFiles.Select(file => new UnifiedMultiYFileState
                {
                    Name = file.Name,
                    FilePath = file.FilePath,
                    Delimiter = file.Delimiter ?? "\t",
                    HeaderRowNumber = file.HeaderRowNumber,
                    Mode = _inputModesByFile.TryGetValue(file.FilePath, out MultiYInputMode mode) ? mode : MultiYInputMode.HeaderMultiY,
                    RawData = GraphReportStorageHelper.CaptureRawTableData(file.FullData),
                    SavedColumnLimits = file.SavedColumnLimits.ToDictionary(pair => pair.Key, pair => new ColumnLimitSetting
                    {
                        ColumnName = pair.Value.ColumnName,
                        SpecValue = pair.Value.SpecValue,
                        UpperValue = pair.Value.UpperValue,
                        LowerValue = pair.Value.LowerValue
                    }, StringComparer.OrdinalIgnoreCase),
                    SelectedColumns = _selectedColumnsByFile.TryGetValue(file.FilePath, out HashSet<string>? selected)
                        ? selected.ToList()
                        : new List<string>()
                }).ToList()
            };

            GraphReportFileDialogHelper.SaveState("Save Graph Report", "unified-multiy.graphreport.json", state);
        }

        private void LoadReportButton_Click(object sender, RoutedEventArgs e)
        {
            UnifiedMultiYReportState? state = GraphReportFileDialogHelper.LoadState<UnifiedMultiYReportState>("Load Graph Report");
            if (state == null)
            {
                MessageBox.Show("Invalid report file.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _loadedFiles.Clear();
            _selectedColumnsByFile.Clear();
            _inputModesByFile.Clear();
            _columnOptions.Clear();
            _currentFile = null;

            foreach (UnifiedMultiYFileState fileState in state.Files)
            {
                var file = GraphMakerFileViewHelper.CreateFileInfo(fileState.FilePath, fileState.Delimiter, fileState.HeaderRowNumber);
                file.Name = fileState.Name;
                file.FullData = GraphReportStorageHelper.BuildTableFromRawData(fileState.RawData);
                file.SavedColumnLimits = fileState.SavedColumnLimits.ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase);
                _loadedFiles.Add(file);
                _inputModesByFile[file.FilePath] = fileState.Mode;
                _selectedColumnsByFile[file.FilePath] = new HashSet<string>(fileState.SelectedColumns, StringComparer.OrdinalIgnoreCase);
            }

            SpecValueTextBox.Text = state.SpecText;
            UpperLimitValueTextBox.Text = state.UpperText;
            LowerLimitValueTextBox.Text = state.LowerText;
            ApplyToAllColumnsCheckBox.IsChecked = state.ApplyToAllColumns;

            FileInfo_DailySampling? selectedFile = _loadedFiles.FirstOrDefault(file =>
                string.Equals(file.FilePath, state.SelectedFilePath, StringComparison.OrdinalIgnoreCase))
                ?? _loadedFiles.FirstOrDefault();
            FileListBox.SelectedItem = selectedFile;
        }

        private void DropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                return;
            }

            e.Effects = DragDropEffects.Copy;
            DropZoneText.Foreground = new SolidColorBrush(Colors.DodgerBlue);
        }

        private void DropZone_DragLeave(object sender, DragEventArgs e)
        {
            DropZoneText.Foreground = (System.Windows.Media.Brush)FindResource("TextSecondaryBrush");
        }

        private void DropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
        }

        private FileInfo_DailySampling BuildCombinedFile(FileInfo_DailySampling sourceFile, IReadOnlyCollection<string> selectedColumns)
        {
            var combinedTable = new DataTable();
            combinedTable.Columns.Add("Date");
            combinedTable.Columns.Add("Value");

            foreach (DataRow row in sourceFile.FullData!.Rows)
            {
                string dateText = row["Date"]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dateText))
                {
                    continue;
                }

                foreach (string columnName in selectedColumns)
                {
                    string valueText = row[columnName]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(valueText))
                    {
                        continue;
                    }

                    DataRow newRow = combinedTable.NewRow();
                    newRow["Date"] = dateText;
                    newRow["Value"] = valueText;
                    combinedTable.Rows.Add(newRow);
                }
            }

            var combinedFile = new FileInfo_DailySampling
            {
                Name = sourceFile.Name,
                FilePath = sourceFile.FilePath,
                FullData = combinedTable,
                Delimiter = sourceFile.Delimiter,
                HeaderRowNumber = 0
            };

            combinedFile.SavedColumnLimits["Value"] = new ColumnLimitSetting
            {
                ColumnName = "Value",
                SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty,
                UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty,
                LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty
            };

            return combinedFile;
        }

        private FileInfo_DailySampling BuildCombinedFile(
            FileInfo_DailySampling sourceFile,
            string xAxisColumnName,
            IReadOnlyCollection<string> selectedColumns,
            bool xAxisIsDate)
        {
            var combinedTable = new DataTable();
            combinedTable.Columns.Add(xAxisColumnName);
            combinedTable.Columns.Add("Value");

            foreach (DataRow row in sourceFile.FullData!.Rows)
            {
                string xAxisText = row[xAxisColumnName]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(xAxisText))
                {
                    continue;
                }

                foreach (string columnName in selectedColumns)
                {
                    string valueText = row[columnName]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(valueText))
                    {
                        continue;
                    }

                    DataRow newRow = combinedTable.NewRow();
                    newRow[0] = xAxisText;
                    newRow[1] = valueText;
                    combinedTable.Rows.Add(newRow);
                }
            }

            var combinedFile = new FileInfo_DailySampling
            {
                Name = sourceFile.Name,
                FilePath = sourceFile.FilePath,
                FullData = combinedTable,
                Delimiter = sourceFile.Delimiter,
                HeaderRowNumber = 0,
                IsXAxisDate = xAxisIsDate,
                ForcedXAxisType = xAxisIsDate ? XAxisValueType.Date : XAxisValueType.Category
            };

            combinedFile.SavedColumnLimits["Value"] = new ColumnLimitSetting
            {
                ColumnName = "Value",
                SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty,
                UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty,
                LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty
            };

            return combinedFile;
        }

        private static DataTable BuildSingleYTable(FileInfo_DailySampling sourceFile, IReadOnlyCollection<string> yColumns)
        {
            var table = new DataTable();
            table.Columns.Add("Date");
            table.Columns.Add("Value");

            foreach (DataRow row in sourceFile.FullData!.Rows)
            {
                string dateText = row[0]?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(dateText))
                {
                    continue;
                }

                foreach (string yColumn in yColumns)
                {
                    string valueText = row[yColumn]?.ToString()?.Trim() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(valueText))
                    {
                        continue;
                    }

                    DataRow newRow = table.NewRow();
                    newRow["Date"] = dateText;
                    newRow["Value"] = valueText;
                    table.Rows.Add(newRow);
                }
            }

            return table;
        }

        private static bool ValidateFirstColumnDate(DataTable table, out string? errorMessage)
        {
            errorMessage = null;
            if (table.Columns.Count == 0)
            {
                errorMessage = "SingleX(Date) - SingleY requires the first column to contain date values.";
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

                if (DateTime.TryParse(text, out _))
                {
                    continue;
                }

                errorMessage = $"First column must contain date values. Invalid value '{text}' was found at row {rowIndex + 1:N0} in column '{firstColumnName}'.";
                return false;
            }

            return true;
        }

        public DateTime? ParseDateString(string dateStr)
        {
            return GraphMakerParsingHelper.TryParseDate(dateStr, out DateTime parsedDate)
                ? parsedDate
                : null;
        }

        public object GetWebModuleSnapshot()
        {
            string delimiter = TabDelimiterRadio.IsChecked == true
                ? "tab"
                : CommaDelimiterRadio.IsChecked == true
                    ? "comma"
                    : "space";

            string currentMode = HeaderMultiYRadio.IsChecked == true ? "Multi Y (Header)"
                : DateSingleYRadio.IsChecked == true ? "Date → Single Y"
                : "Date → Multi Y (No Header)";

            return new
            {
                moduleType = "GraphMakerUnifiedMultiY",
                currentFileName = _currentFile?.Name ?? "(Select a file)",
                currentMode,
                modeHint = ModeHintTextBlock.Text ?? string.Empty,
                delimiter,
                headerRow = HeaderRowTextBox.Text ?? "1",
                applyToAllColumns = ApplyToAllColumnsCheckBox.IsChecked == true,
                combinedYAxisView = _combinedYAxisView,
                combinedXAxisColumn = _combinedXAxisColumn,
                combinedXAxisType = _combinedXAxisIsDate ? "date" : "text",
                specValue = SpecValueTextBox.Text ?? string.Empty,
                upperValue = UpperLimitValueTextBox.Text ?? string.Empty,
                lowerValue = LowerLimitValueTextBox.Text ?? string.Empty,
                files = _loadedFiles.Select(file => new
                {
                    name = file.Name,
                    filePath = file.FilePath,
                    isSelected = string.Equals(file.FilePath, _currentFile?.FilePath, StringComparison.OrdinalIgnoreCase)
                }).ToArray(),
                columnOptions = _columnOptions.Select(option => new
                {
                    name = option.ColumnName,
                    isSelected = option.IsSelected
                }).ToArray(),
                combinedXAxisOptions = _columnOptions.Select(option => option.ColumnName).ToArray(),
                previewColumns = BuildPreviewColumns(_currentFile?.FullData),
                previewRows = BuildPreviewRows(_currentFile?.FullData)
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("inputMode", out JsonElement inputModeElement) &&
                Enum.TryParse(inputModeElement.GetString(), out MultiYInputMode inputMode) &&
                IsLoaded)
            {
                _isApplyingModeUi = true;
                HeaderMultiYRadio.IsChecked = inputMode == MultiYInputMode.HeaderMultiY;
                DateSingleYRadio.IsChecked = inputMode == MultiYInputMode.DateSingleY;
                DateNoHeaderMultiYRadio.IsChecked = inputMode == MultiYInputMode.DateNoHeaderMultiY;
                _isApplyingModeUi = false;
                if (_currentFile != null)
                {
                    _inputModesByFile[_currentFile.FilePath] = inputMode;
                    _currentFile.FullData = null!;
                    LoadAndBindCurrentFile();
                }
            }

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

            if (payload.TryGetProperty("applyToAllColumns", out JsonElement applyElement))
            {
                ApplyToAllColumnsCheckBox.IsChecked = applyElement.GetBoolean();
            }

            if (payload.TryGetProperty("combinedYAxisView", out JsonElement combinedYAxisViewElement))
            {
                _combinedYAxisView = combinedYAxisViewElement.GetBoolean();
            }

            if (payload.TryGetProperty("combinedXAxisColumn", out JsonElement combinedXAxisColumnElement))
            {
                _combinedXAxisColumn = combinedXAxisColumnElement.GetString() ?? string.Empty;
            }

            if (payload.TryGetProperty("combinedXAxisType", out JsonElement combinedXAxisTypeElement))
            {
                _combinedXAxisIsDate = !string.Equals(combinedXAxisTypeElement.GetString(), "text", StringComparison.OrdinalIgnoreCase);
            }

            if (payload.TryGetProperty("specValue", out JsonElement specElement))
            {
                SpecValueTextBox.Text = specElement.GetString() ?? string.Empty;
            }

            if (payload.TryGetProperty("upperValue", out JsonElement upperElement))
            {
                UpperLimitValueTextBox.Text = upperElement.GetString() ?? string.Empty;
            }

            if (payload.TryGetProperty("lowerValue", out JsonElement lowerElement))
            {
                LowerLimitValueTextBox.Text = lowerElement.GetString() ?? string.Empty;
            }

            if (payload.TryGetProperty("selectedFilePath", out JsonElement selectedFileElement))
            {
                string? selectedFilePath = selectedFileElement.GetString();
                FileInfo_DailySampling? selected = _loadedFiles.FirstOrDefault(file => string.Equals(file.FilePath, selectedFilePath, StringComparison.OrdinalIgnoreCase));
                if (selected is not null)
                {
                    FileListBox.SelectedItem = selected;
                }
            }

            if (payload.TryGetProperty("selectedColumns", out JsonElement selectedColumnsElement) &&
                selectedColumnsElement.ValueKind == JsonValueKind.Array &&
                _currentFile is not null)
            {
                HashSet<string> selectedNames = selectedColumnsElement.EnumerateArray()
                    .Select(item => item.GetString() ?? string.Empty)
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                foreach (SelectableColumnOption option in _columnOptions)
                {
                    option.IsSelected = selectedNames.Contains(option.ColumnName);
                }

                _selectedColumnsByFile[_currentFile.FilePath] = selectedNames;
                if (string.IsNullOrWhiteSpace(_combinedXAxisColumn) ||
                    !selectedNames.Contains(_combinedXAxisColumn))
                {
                    _combinedXAxisColumn = selectedNames.FirstOrDefault() ?? _columnOptions.FirstOrDefault()?.ColumnName ?? string.Empty;
                }
                BindPreviewGrid(_currentFile);
            }

            NotifyWebModuleSnapshotChanged();
            return GetWebModuleSnapshot();
        }

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "browse-files":
                    BrowseButton_Click(this, new RoutedEventArgs());
                    break;
                case "load-report":
                    LoadReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "save-report":
                    SaveReportButton_Click(this, new RoutedEventArgs());
                    break;
                case "apply-limits":
                    ApplyLimitsButton_Click(this, new RoutedEventArgs());
                    break;
                case "generate-graph":
                    GenerateGraphButton_Click(this, new RoutedEventArgs());
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
