using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using UserControl = System.Windows.Controls.UserControl;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;

namespace GraphMaker
{
    public partial class ValuePlotMultiColumnView : GraphViewBase
    {
        private readonly ObservableCollection<FileInfo_DailySampling> _loadedFiles = new();
        private readonly ObservableCollection<SelectableColumnOption> _columnOptions = new();
        private readonly Dictionary<string, HashSet<string>> _selectedColumnsByFile = new(StringComparer.OrdinalIgnoreCase);

        private FileInfo_DailySampling? _currentFile;
        private DataTable? _previewTable;
        private bool _isRefreshingPreview;
        private bool _combinedYAxisView;
        private string _combinedXAxisColumn = string.Empty;
        private bool _combinedXAxisIsDate = true;
        public ValuePlotMultiColumnView()
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
                LoadFile(filePath);
            }
        }

        private void LoadFile(string filePath)
        {
            if (_loadedFiles.Any(f => string.Equals(f.FilePath, filePath, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            var file = GraphMakerFileViewHelper.CreateFileInfo(filePath, "\t", 1);

            _loadedFiles.Add(file);
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
                CurrentFileNameText.Text = "(Select a file)";
                RowCountText.Text = "0";
                ColumnCountText.Text = "0";
                return;
            }

            _currentFile.Delimiter = _currentFile.Delimiter ?? "\t";
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
                ParseFileIntoDataTable(_currentFile);
                BuildPreviewTable(_currentFile);
                BindColumnOptions(_currentFile);

                CurrentFileNameText.Text = _currentFile.Name;
                RowCountText.Text = _currentFile.FullData?.Rows.Count.ToString("N0") ?? "0";
                ColumnCountText.Text = _currentFile.FullData?.Columns.Count.ToString() ?? "0";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ParseFileIntoDataTable(FileInfo_DailySampling file)
        {
            GraphMakerFileViewHelper.ParseHeaderIntoDataTable(file, 2);
            DataTable table = file.FullData!;
            var dates = table.Rows.Cast<DataRow>()
                .Select(row => row[0]?.ToString() ?? string.Empty)
                .ToList();
            var sampleNumbers = table.Columns.Cast<DataColumn>()
                .Skip(1)
                .Select(column => column.ColumnName)
                .ToList();
            file.Dates = dates;
            file.SampleNumbers = sampleNumbers;
            GraphMakerFileViewHelper.EnsureColumnLimits(
                file,
                name => !string.Equals(name, table.Columns[0].ColumnName, StringComparison.OrdinalIgnoreCase));
        }

        private void BuildPreviewTable(FileInfo_DailySampling file)
        {
            if (file.FullData == null)
            {
                _previewTable = null;
                DataPreviewGrid.ItemsSource = null;
                return;
            }

            var preview = file.FullData.Clone();

            DataRow upper = preview.NewRow();
            DataRow spec = preview.NewRow();
            DataRow lower = preview.NewRow();

            upper[0] = "Upper Limit";
            spec[0] = "Spec";
            lower[0] = "Lower Limit";

            for (int c = 1; c < preview.Columns.Count; c++)
            {
                string colName = preview.Columns[c].ColumnName;
                if (file.SavedColumnLimits.TryGetValue(colName, out var saved))
                {
                    upper[c] = saved.UpperValue;
                    spec[c] = saved.SpecValue;
                    lower[c] = saved.LowerValue;
                }
            }

            preview.Rows.Add(upper);
            preview.Rows.Add(spec);
            preview.Rows.Add(lower);

            foreach (DataRow src in file.FullData.Rows)
            {
                DataRow dst = preview.NewRow();
                dst.ItemArray = (object[])src.ItemArray.Clone();
                preview.Rows.Add(dst);
            }

            _previewTable = preview;
            _isRefreshingPreview = true;
            DataPreviewGrid.ItemsSource = preview.DefaultView;
            _isRefreshingPreview = false;
        }

        private void BindColumnOptions(FileInfo_DailySampling file)
        {
            if (file.FullData == null || file.FullData.Columns.Count == 0)
            {
                _columnOptions.Clear();
                return;
            }

            string xColumnName = file.FullData.Columns[0].ColumnName;
            GraphMakerFileViewHelper.BindColumnOptions(
                file,
                _columnOptions,
                _selectedColumnsByFile,
                name => !string.Equals(name, xColumnName, StringComparison.OrdinalIgnoreCase));

            if (_columnOptions.Count > 0)
            {
                var first = _columnOptions[0].ColumnName;
                if (file.SavedColumnLimits.TryGetValue(first, out var firstLimit))
                {
                    SpecValueTextBox.Text = firstLimit.SpecValue;
                    UpperLimitValueTextBox.Text = firstLimit.UpperValue;
                    LowerLimitValueTextBox.Text = firstLimit.LowerValue;
                }
            }

            if (string.IsNullOrWhiteSpace(_combinedXAxisColumn) ||
                !_columnOptions.Any(option => string.Equals(option.ColumnName, _combinedXAxisColumn, StringComparison.OrdinalIgnoreCase)))
            {
                _combinedXAxisColumn = _columnOptions.FirstOrDefault()?.ColumnName ?? string.Empty;
            }
        }

        private void SaveCurrentSelectionState()
        {
            GraphMakerFileViewHelper.SaveCurrentSelectionState(_currentFile, _columnOptions, _selectedColumnsByFile);
        }

        private string GetCurrentDelimiter()
        {
            if (CommaDelimiterRadio.IsChecked == true)
            {
                return ",";
            }

            if (SpaceDelimiterRadio.IsChecked == true)
            {
                return " ";
            }

            return "\t";
        }

        private int ParseHeaderRowNumber()
        {
            if (int.TryParse(HeaderRowTextBox.Text, out int row) && row > 0)
            {
                return row;
            }

            return 1;
        }

        private void RemoveFileButton_Click(object sender, RoutedEventArgs e)
        {
            if (FileListBox.SelectedItem is not FileInfo_DailySampling selected)
            {
                return;
            }

            _loadedFiles.Remove(selected);
            _selectedColumnsByFile.Remove(selected.FilePath);

            if (_loadedFiles.Count == 0)
            {
                _currentFile = null;
                _previewTable = null;
                DataPreviewGrid.ItemsSource = null;
                _columnOptions.Clear();
                CurrentFileNameText.Text = "(Select a file)";
                RowCountText.Text = "0";
                ColumnCountText.Text = "0";
                return;
            }

            FileListBox.SelectedIndex = 0;
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

            _currentFile.FullData = null;
            _currentFile.Dates = null;
            _currentFile.SampleNumbers = null;
            LoadAndBindCurrentFile();
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
            _currentFile.FullData = null;
            _currentFile.Dates = null;
            _currentFile.SampleNumbers = null;
            LoadAndBindCurrentFile();
        }

        private void ApplyLimitsToAllColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile?.FullData == null)
            {
                return;
            }

            string spec = SpecValueTextBox.Text?.Trim() ?? string.Empty;
            string upper = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty;
            string lower = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty;

            if (ApplyToAllColumnsCheckBox.IsChecked == true)
            {
                for (int c = 1; c < _currentFile.FullData.Columns.Count; c++)
                {
                    string colName = _currentFile.FullData.Columns[c].ColumnName;
                    if (!_currentFile.SavedColumnLimits.TryGetValue(colName, out var limit))
                    {
                        limit = new ColumnLimitSetting { ColumnName = colName };
                        _currentFile.SavedColumnLimits[colName] = limit;
                    }

                    limit.SpecValue = spec;
                    limit.UpperValue = upper;
                    limit.LowerValue = lower;
                }
            }
            else
            {
                foreach (var option in _columnOptions.Where(x => x.IsSelected))
                {
                    if (!_currentFile.SavedColumnLimits.TryGetValue(option.ColumnName, out var limit))
                    {
                        limit = new ColumnLimitSetting { ColumnName = option.ColumnName };
                        _currentFile.SavedColumnLimits[option.ColumnName] = limit;
                    }

                    limit.SpecValue = spec;
                    limit.UpperValue = upper;
                    limit.LowerValue = lower;
                }
            }

            BuildPreviewTable(_currentFile);
        }

        private void DataPreviewGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            if (e.PropertyName == "RowError")
            {
                e.Cancel = true;
                return;
            }

            if (e.Column is DataGridTextColumn textColumn)
            {
                textColumn.IsReadOnly = false;
            }

            e.Column.Width = new DataGridLength(1, DataGridLengthUnitType.Star);
        }

        private void DataPreviewGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            if (_isRefreshingPreview)
            {
                return;
            }

            if (_previewTable == null)
            {
                e.Cancel = true;
                return;
            }

            if (e.Row.GetIndex() > 2)
            {
                e.Cancel = true;
                return;
            }

            if (e.Column.DisplayIndex == 0)
            {
                e.Cancel = true;
            }
        }

        private void DataPreviewGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (_isRefreshingPreview || _currentFile?.FullData == null || _previewTable == null)
            {
                return;
            }

            int rowIndex = e.Row.GetIndex();
            if (rowIndex < 0 || rowIndex > 2)
            {
                return;
            }

            if (e.Column.DisplayIndex <= 0)
            {
                return;
            }

            if (e.EditingElement is not TextBox tb)
            {
                return;
            }

            string colName = _previewTable.Columns[e.Column.DisplayIndex].ColumnName;
            string value = tb.Text?.Trim() ?? string.Empty;

            if (!_currentFile.SavedColumnLimits.TryGetValue(colName, out var limit))
            {
                limit = new ColumnLimitSetting { ColumnName = colName };
                _currentFile.SavedColumnLimits[colName] = limit;
            }

            switch (rowIndex)
            {
                case 0:
                    limit.UpperValue = value;
                    break;
                case 1:
                    limit.SpecValue = value;
                    break;
                case 2:
                    limit.LowerValue = value;
                    break;
            }
        }

        private void DataPreviewGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control || e.Key != Key.V)
            {
                return;
            }

            e.Handled = true;
            PasteLimitsFromClipboard();
        }

        private void PasteLimitsFromClipboard()
        {
            if (_previewTable == null || _currentFile?.FullData == null)
            {
                return;
            }

            if (!Clipboard.ContainsText())
            {
                return;
            }

            string text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            int startRow = DataPreviewGrid.Items.IndexOf(DataPreviewGrid.CurrentItem);
            int startColumn = DataPreviewGrid.CurrentColumn?.DisplayIndex ?? 1;

            if (startRow < 0)
            {
                startRow = 0;
            }

            if (startColumn <= 0)
            {
                startColumn = 1;
            }

            if (startRow > 2 || startColumn >= _previewTable.Columns.Count)
            {
                return;
            }

            string[] rowTokens = text
                .Replace("\r\n", "\n")
                .Replace('\r', '\n')
                .Split('\n', StringSplitOptions.RemoveEmptyEntries);

            if (rowTokens.Length == 0)
            {
                return;
            }

            _isRefreshingPreview = true;
            for (int r = 0; r < rowTokens.Length; r++)
            {
                int targetRow = startRow + r;
                if (targetRow > 2)
                {
                    break;
                }

                string[] cells = rowTokens[r].Split('\t');
                for (int c = 0; c < cells.Length; c++)
                {
                    int targetColumn = startColumn + c;
                    if (targetColumn <= 0 || targetColumn >= _previewTable.Columns.Count)
                    {
                        continue;
                    }

                    string value = cells[c].Trim();
                    _previewTable.Rows[targetRow][targetColumn] = value;
                    UpdateSavedColumnLimit(targetRow, targetColumn, value);
                }
            }

            _isRefreshingPreview = false;
            DataPreviewGrid.Items.Refresh();
        }

        private void UpdateSavedColumnLimit(int rowIndex, int columnIndex, string value)
        {
            if (_previewTable == null || _currentFile == null || columnIndex <= 0)
            {
                return;
            }

            string colName = _previewTable.Columns[columnIndex].ColumnName;
            if (!_currentFile.SavedColumnLimits.TryGetValue(colName, out var limit))
            {
                limit = new ColumnLimitSetting { ColumnName = colName };
                _currentFile.SavedColumnLimits[colName] = limit;
            }

            switch (rowIndex)
            {
                case 0:
                    limit.UpperValue = value;
                    break;
                case 1:
                    limit.SpecValue = value;
                    break;
                case 2:
                    limit.LowerValue = value;
                    break;
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

            var selectedColumns = _selectedColumnsByFile.TryGetValue(_currentFile.FilePath, out var selectedSet)
                ? selectedSet.ToList()
                : _currentFile.FullData.Columns.Cast<DataColumn>().Skip(1).Select(c => c.ColumnName).ToList();
            if (selectedColumns.Count == 0)
            {
                MessageBox.Show("Select at least one data column to plot.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            FileInfo_DailySampling sourceFile = _currentFile;
            List<string> columnsToPlot = selectedColumns.ToList();

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

                sourceFile = BuildCombinedFile(_currentFile, _combinedXAxisColumn, combinedYColumns, _combinedXAxisIsDate);
                columnsToPlot = new List<string> { "Value" };
            }

            var result = MultiColumnGraphCalculator.Calculate(sourceFile, columnsToPlot);
            string xAxisTitle = _combinedYAxisView
                ? _combinedXAxisColumn
                : _currentFile.FullData.Columns[0].ColumnName;
            var window = new MultiColumnResultWindow(_currentFile.Name, xAxisTitle, result)
            {
                Owner = Window.GetWindow(this),
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowState = WindowState.Maximized
            };
            window.Show();
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
                HeaderRowNumber = sourceFile.HeaderRowNumber,
                IsXAxisDate = xAxisIsDate
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
                    LoadFile(file);
                }
            }
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

        public DateTime? ParseDateString(string dateStr)
        {
            if (DateTime.TryParse(dateStr, out DateTime dt))
            {
                return dt;
            }

            return null;
        }

        private void OpenFileSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentFile == null)
            {
                MessageBox.Show("Select a file first.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            LoadAndBindCurrentFile();
        }

        public object GetWebModuleSnapshot()
        {
            string delimiter = TabDelimiterRadio.IsChecked == true ? "tab"
                : CommaDelimiterRadio.IsChecked == true ? "comma"
                : "space";

            return new
            {
                moduleType = "GraphMakerUnifiedMultiY",
                currentFileName = CurrentFileNameText.Text ?? "(Select a file)",
                rowCount = RowCountText.Text ?? "0",
                columnCount = ColumnCountText.Text ?? "0",
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
                selectedColumns = _columnOptions.Where(option => option.IsSelected).Select(option => option.ColumnName).ToArray(),
                columnOptions = _columnOptions.Select(option => new
                {
                    name = option.ColumnName,
                    isSelected = option.IsSelected
                }).ToArray(),
                combinedXAxisOptions = _columnOptions.Select(option => option.ColumnName).ToArray()
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
            }

            if (payload.TryGetProperty("headerRow", out JsonElement headerRowElement))
            {
                HeaderRowTextBox.Text = headerRowElement.GetString() ?? "1";
            }

            if (payload.TryGetProperty("applyToAllColumns", out JsonElement applyToAllElement))
            {
                ApplyToAllColumnsCheckBox.IsChecked = applyToAllElement.GetBoolean();
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

            if (payload.TryGetProperty("selectedFilePath", out JsonElement selectedFilePathElement))
            {
                string? selectedFilePath = selectedFilePathElement.GetString();
                FileInfo_DailySampling? selectedFile = _loadedFiles.FirstOrDefault(file => string.Equals(file.FilePath, selectedFilePath, StringComparison.OrdinalIgnoreCase));
                if (selectedFile is not null)
                {
                    FileListBox.SelectedItem = selectedFile;
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
                case "remove-selected":
                    RemoveFileButton_Click(this, new RoutedEventArgs());
                    break;
                case "apply-limits":
                    ApplyLimitsToAllColumnsButton_Click(this, new RoutedEventArgs());
                    break;
                case "generate-graph":
                    GenerateGraphButton_Click(this, new RoutedEventArgs());
                    break;
            }

            return GetWebModuleSnapshot();
        }

    }
}
