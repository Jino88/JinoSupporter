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
using Microsoft.Win32;
using DataFormats = System.Windows.DataFormats;
using DragDropEffects = System.Windows.DragDropEffects;
using DragEventArgs = System.Windows.DragEventArgs;
using MessageBox = System.Windows.MessageBox;
using UserControl = System.Windows.Controls.UserControl;

namespace GraphMaker
{
    public partial class DateNoHeaderMultiYView : UserControl
    {
        private readonly ObservableCollection<FileInfo_DailySampling> _loadedFiles = new();
        private readonly ObservableCollection<SelectableColumnOption> _columnOptions = new();
        private readonly Dictionary<string, HashSet<string>> _selectedColumnsByFile = new(StringComparer.OrdinalIgnoreCase);
        private FileInfo_DailySampling? _currentFile;

        public DateNoHeaderMultiYView()
        {
            InitializeComponent();
            FileListBox.ItemsSource = _loadedFiles;
            ColumnOptionListBox.ItemsSource = _columnOptions;
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

            var file = GraphMakerFileViewHelper.CreateFileInfo(filePath, "\t", 0);

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
                GraphMakerFileViewHelper.ParseDateNoHeaderIntoDataTable(_currentFile);
                GraphMakerFileViewHelper.EnsureColumnLimits(
                    _currentFile,
                    name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase));
                DataPreviewGrid.ItemsSource = _currentFile.FullData?.DefaultView;
                BindColumnOptions(_currentFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BindColumnOptions(FileInfo_DailySampling file)
        {
            GraphMakerFileViewHelper.BindColumnOptions(
                file,
                _columnOptions,
                _selectedColumnsByFile,
                name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase));
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
                ? _currentFile.FullData.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase))
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
                : _currentFile.FullData.Columns.Cast<DataColumn>().Select(c => c.ColumnName).Where(name => !string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase)).ToList();

            if (selectedColumns.Count == 0)
            {
                MessageBox.Show("Select at least one data column to plot.", "Notice", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            FileInfo_DailySampling combinedFile = BuildCombinedFile(_currentFile, selectedColumns);
            var result = MultiColumnGraphCalculator.Calculate(combinedFile, new List<string> { "Value" });
            var window = new MultiColumnResultWindow(_currentFile.Name, "Date", result)
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

            var combinedLimit = new ColumnLimitSetting
            {
                ColumnName = "Value",
                SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty,
                UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty,
                LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty
            };
            combinedFile.SavedColumnLimits["Value"] = combinedLimit;
            return combinedFile;
        }
    }
}
