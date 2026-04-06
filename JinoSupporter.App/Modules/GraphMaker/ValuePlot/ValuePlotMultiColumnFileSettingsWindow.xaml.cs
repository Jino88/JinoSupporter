using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace GraphMaker
{
    public partial class ValuePlotMultiColumnFileSettingsWindow : Window
    {
        private sealed class SelectableColumn : INotifyPropertyChanged
        {
            private bool _isSelected;

            public string Name { get; init; } = string.Empty;

            public bool IsSelected
            {
                get => _isSelected;
                set
                {
                    if (_isSelected == value)
                    {
                        return;
                    }

                    _isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }

            public event PropertyChangedEventHandler? PropertyChanged;
        }

        private readonly string _filePath;
        private readonly List<SelectableColumn> _columns = new();
        private bool _isLoadingColumns;

        public string Delimiter { get; private set; } = "\t";
        public int HeaderRowNumber { get; private set; } = 1;
        public string SpecValue { get; private set; } = string.Empty;
        public string UpperValue { get; private set; } = string.Empty;
        public string LowerValue { get; private set; } = string.Empty;
        public bool ApplyToAllColumns { get; private set; } = true;
        public IReadOnlyList<string> SelectedColumns { get; private set; } = Array.Empty<string>();
        public IReadOnlyList<string> AllDataColumns { get; private set; } = Array.Empty<string>();

        public ValuePlotMultiColumnFileSettingsWindow(
            string filePath,
            FileInfo_DailySampling fileInfo,
            IEnumerable<string>? selectedColumns)
        {
            InitializeComponent();
            _filePath = filePath;

            if (fileInfo.Delimiter == ",")
            {
                CommaDelimiterRadio.IsChecked = true;
            }
            else if (fileInfo.Delimiter == " ")
            {
                SpaceDelimiterRadio.IsChecked = true;
            }
            else
            {
                TabDelimiterRadio.IsChecked = true;
            }

            HeaderRowTextBox.Text = fileInfo.HeaderRowNumber.ToString();
            ApplyToAllColumnsCheckBox.IsChecked = true;

            SpecValueTextBox.Text = fileInfo.SavedSpecValue ?? string.Empty;
            UpperLimitValueTextBox.Text = fileInfo.SavedUpperValue ?? string.Empty;
            LowerLimitValueTextBox.Text = fileInfo.SavedLowerValue ?? string.Empty;

            LoadColumns(selectedColumns);
            RefreshPreview();
        }

        private void ReloadColumnsButton_Click(object sender, RoutedEventArgs e)
        {
            LoadColumns(_columns.Where(c => c.IsSelected).Select(c => c.Name), showValidationErrors: true);
        }

        private void LoadColumns(IEnumerable<string>? selectedColumns, bool showValidationErrors = true)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                if (showValidationErrors)
                {
                    MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                PreviewDataGrid.ItemsSource = null;
                PreviewSummaryTextBlock.Text = "Header row must be a positive integer.";
                return;
            }

            var delimiter = GetDelimiterFromUi();
            var headers = ReadHeaders(_filePath, delimiter, headerRow);
            if (headers.Count <= 1)
            {
                if (showValidationErrors)
                {
                    MessageBox.Show("Need at least two columns (X + data columns).", "Invalid Data",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                PreviewDataGrid.ItemsSource = null;
                PreviewSummaryTextBlock.Text = "Need at least two columns (X + data columns).";
                return;
            }

            var selectedSet = new HashSet<string>(selectedColumns ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            var dataColumns = headers.Skip(1).ToList();
            _isLoadingColumns = true;
            _columns.Clear();
            foreach (var name in dataColumns)
            {
                var column = new SelectableColumn
                {
                    Name = name,
                    IsSelected = selectedSet.Count == 0 || selectedSet.Contains(name)
                };
                column.PropertyChanged += Column_PropertyChanged;
                _columns.Add(column);
            }

            ColumnListBox.ItemsSource = null;
            ColumnListBox.ItemsSource = _columns;
            AllDataColumns = dataColumns;
            _isLoadingColumns = false;

            RefreshPreview();
        }

        private static List<string> ReadHeaders(string filePath, string delimiter, int headerRowNumber)
        {
            var lines = File.ReadAllLines(filePath);
            var headerIndex = headerRowNumber - 1;
            if (headerIndex < 0 || headerIndex >= lines.Length)
            {
                return new List<string>();
            }

            var rawHeaders = GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter);
            return GraphMakerTableHelper.BuildUniqueHeaders(rawHeaders);
        }

        private string GetDelimiterFromUi()
        {
            return TabDelimiterRadio.IsChecked == true ? "\t" :
                CommaDelimiterRadio.IsChecked == true ? "," : " ";
        }

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            _isLoadingColumns = true;
            foreach (var item in _columns)
            {
                item.IsSelected = true;
            }
            _isLoadingColumns = false;
            RefreshPreview();
        }

        private void ClearAllButton_Click(object sender, RoutedEventArgs e)
        {
            _isLoadingColumns = true;
            foreach (var item in _columns)
            {
                item.IsSelected = false;
            }
            _isLoadingColumns = false;
            RefreshPreview();
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                MessageBox.Show("Header row must be a positive integer.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var selected = _columns.Where(c => c.IsSelected).Select(c => c.Name).ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show("Select at least one column to plot.", "Invalid Setting",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            Delimiter = GetDelimiterFromUi();
            HeaderRowNumber = headerRow;
            SpecValue = SpecValueTextBox.Text?.Trim() ?? string.Empty;
            UpperValue = UpperLimitValueTextBox.Text?.Trim() ?? string.Empty;
            LowerValue = LowerLimitValueTextBox.Text?.Trim() ?? string.Empty;
            ApplyToAllColumns = ApplyToAllColumnsCheckBox.IsChecked == true;
            SelectedColumns = selected;

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void Column_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_isLoadingColumns || e.PropertyName != nameof(SelectableColumn.IsSelected))
            {
                return;
            }

            RefreshPreview();
        }

        private void PreviewSettingChanged(object sender, RoutedEventArgs e)
        {
            RefreshPreviewWithReload();
        }

        private void HeaderRowTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            RefreshPreviewWithReload();
        }

        private void RefreshPreviewWithReload()
        {
            LoadColumns(_columns.Where(c => c.IsSelected).Select(c => c.Name), showValidationErrors: false);
        }

        private void RefreshPreview()
        {
            if (!int.TryParse(HeaderRowTextBox.Text, out var headerRow) || headerRow <= 0)
            {
                PreviewDataGrid.ItemsSource = null;
                PreviewSummaryTextBlock.Text = "Header row must be a positive integer.";
                return;
            }

            try
            {
                var delimiter = GetDelimiterFromUi();
                var lines = File.ReadAllLines(_filePath);
                var headerIndex = headerRow - 1;
                if (headerIndex < 0 || headerIndex >= lines.Length)
                {
                    PreviewDataGrid.ItemsSource = null;
                    PreviewSummaryTextBlock.Text = "Header row is outside file range.";
                    return;
                }

                var headers = GraphMakerTableHelper.BuildUniqueHeaders(GraphMakerTableHelper.SplitLine(lines[headerIndex], delimiter));
                if (headers.Count == 0)
                {
                    PreviewDataGrid.ItemsSource = null;
                    PreviewSummaryTextBlock.Text = "No columns found.";
                    return;
                }

                var selected = _columns.Where(c => c.IsSelected).Select(c => c.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var previewHeaders = new List<string> { headers[0] };
                previewHeaders.AddRange(headers.Skip(1).Where(h => selected.Count == 0 || selected.Contains(h)));

                var table = new DataTable();
                foreach (var header in previewHeaders)
                {
                    table.Columns.Add(header);
                }

                var maxRows = 12;
                for (var i = headerIndex + 1; i < lines.Length && table.Rows.Count < maxRows; i++)
                {
                    if (string.IsNullOrWhiteSpace(lines[i]))
                    {
                        continue;
                    }

                    var values = GraphMakerTableHelper.SplitLine(lines[i], delimiter);
                    var row = table.NewRow();
                    for (var col = 0; col < previewHeaders.Count; col++)
                    {
                        var sourceIndex = headers.FindIndex(h => h.Equals(previewHeaders[col], StringComparison.OrdinalIgnoreCase));
                        row[col] = sourceIndex >= 0 && sourceIndex < values.Length ? values[sourceIndex] : string.Empty;
                    }
                    table.Rows.Add(row);
                }

                PreviewDataGrid.ItemsSource = table.DefaultView;
                var selectedCount = Math.Max(0, previewHeaders.Count - 1);
                PreviewSummaryTextBlock.Text = $"Showing {table.Rows.Count} rows, X + {selectedCount} data columns.";
            }
            catch (Exception ex)
            {
                PreviewDataGrid.ItemsSource = null;
                PreviewSummaryTextBlock.Text = $"Preview unavailable: {ex.Message}";
            }
        }
    }
}
