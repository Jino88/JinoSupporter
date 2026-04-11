using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using JinoSupporter.Controls;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GraphMaker
{
    public partial class NoXMultiYView : UserControl
    {
        private readonly ObservableCollection<FileInfo_DailySampling> _loadedFiles = new();
        private readonly ObservableCollection<SelectableColumnOption> _columnOptions = new();
        private readonly Dictionary<string, HashSet<string>> _selectedColumnsByFile = new(StringComparer.OrdinalIgnoreCase);

        private FileInfo_DailySampling? _currentFile;

        public NoXMultiYView()
        {
            InitializeComponent();
            FileListBox.ItemsSource = _loadedFiles;
            ColumnOptionListBox.ItemsSource = _columnOptions;
            UpdateWorkflowSummary();
        }

        private void FileDropBox_FilesSelected(object sender, FilesSelectedEventArgs e)
        {
            foreach (var path in e.FilePaths)
            {
                LoadFile(path);
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
                UpdateWorkflowSummary();
                return;
            }

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
                GraphMakerFileViewHelper.ParseHeaderIntoDataTable(_currentFile, 1);
                GraphMakerFileViewHelper.EnsureColumnLimits(_currentFile);
                DataPreviewGrid.ItemsSource = _currentFile.FullData?.DefaultView;
                BindColumnOptions(_currentFile);
                UpdateWorkflowSummary();
            }
            catch (Exception ex)
            {
                StatusText.Text = $"Failed to load file: {ex.Message}";
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BindColumnOptions(FileInfo_DailySampling file)
        {
            GraphMakerFileViewHelper.BindColumnOptions(file, _columnOptions, _selectedColumnsByFile);
        }

        private void SaveCurrentSelectionState()
        {
            GraphMakerFileViewHelper.SaveCurrentSelectionState(_currentFile, _columnOptions, _selectedColumnsByFile);
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
                ? _currentFile.FullData.Columns.Cast<DataColumn>().Select(c => c.ColumnName)
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
                : _currentFile.FullData.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();

            if (selectedColumns.Count == 0)
            {
                MessageBox.Show("Select at least one data column to plot.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = NoXMultiYGraphCalculator.Calculate(_currentFile, selectedColumns);
            var window = new NoXMultiYResultWindow(_currentFile.Name, result)
            {
                Owner = Window.GetWindow(this),
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                WindowState = WindowState.Maximized
            };
            window.Show();
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
                DataPreviewGrid.ItemsSource = null;
                _columnOptions.Clear();
                UpdateWorkflowSummary();
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

            _currentFile.Delimiter = TabDelimiterRadio.IsChecked == true ? "\t"
                : CommaDelimiterRadio.IsChecked == true ? ","
                : " ";

            _currentFile.FullData = null!;
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
            _currentFile.FullData = null!;
            LoadAndBindCurrentFile();
        }

        private void UpdateWorkflowSummary()
        {
            CurrentFileNameText.Text = _currentFile?.Name ?? "(No file)";
            RowCountText.Text = _currentFile?.FullData?.Rows.Count.ToString("N0") ?? "0";
            ColumnCountText.Text = _currentFile?.FullData?.Columns.Count.ToString() ?? "0";
            StatusText.Text = _loadedFiles.Count == 0
                ? "Drop file into the workflow area or click Browse to load data."
                : $"Loaded {_loadedFiles.Count:N0} file(s).";
        }

    }
}
